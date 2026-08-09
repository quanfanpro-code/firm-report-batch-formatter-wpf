using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FirmFormatter.OpenXml.Core;

public sealed class ParagraphService
{
    // 一/二/三级标题正则与 OpenXmlHelper 共用同一份定义，避免两处维护漂移
    private static readonly Regex H1 = OpenXmlHelper.一级标题文本;
    private static readonly Regex H2 = OpenXmlHelper.二级标题文本;
    private static readonly Regex H3 = OpenXmlHelper.三级标题文本;
    private static readonly Regex H3Separator = new(@"^(\s*\d+)\s*[、\.．]\s*", RegexOptions.Compiled);
    private static readonly Regex DocNumberRegex = new(@"^川华信\S{0,8}[（(]\d{4}[)）]第?\d{1,6}号$", RegexOptions.Compiled);
    private static readonly string[] HeaderOrgKeywords = ["公司", "集团", "委员会", "事务所", "研究院", "中心", "学校", "学院", "银行"];

    public void Apply(WordprocessingDocument word, bool hasCover, List<Paragraph> signoffParagraphs)
    {
        var body = word.MainDocumentPart?.Document?.Body;
        if (body is null) return;

        var signoffSet = signoffParagraphs != null ? new HashSet<Paragraph>(signoffParagraphs) : null;

        // 一次性建立 (abstractNumId, ilvl) → 长文本段落 索引，供 ForceNumberingIndent 守卫 O(1) 查询，避免每个编号标题都全篇扫描（O(N²)）
        var 长编号段落索引 = 建立长编号段落索引(word.MainDocumentPart);

        var sectionIndex = 0;
        var inCoverZone = hasCover;
        bool checkingTitle = true;
        int titleLinesFound = 0;
        bool checkingDocNumber = false;
        int docNumberScanCount = 0;
        bool checkingHeaderLines = false;
        int headerLinesFound = 0;

        foreach (var p in body.Descendants<Paragraph>())
        {
            // [核心防御]：绝对不能碰表格里的段落，否则会把 TableService 辛辛苦苦调教的单倍行距和无缩进全部覆盖成 21 磅和 2 字符缩进！
            // 文本框（w:txbxContent）里的段落同样豁免，否则会被强套正文格式，破坏侧栏/签章文本框排版
            if (p.Ancestors<Table>().Any() || p.Ancestors<TextBoxContent>().Any())
            {
                continue;
            }

            // [核心防御2]：绝对不能碰真实的落款区！
            if (signoffSet != null && signoffSet.Contains(p))
            {
                continue;
            }

            处理正文段落(p);

            // [统一出口]：无论上面走了哪个分支（封面区/空段/标题/文号/header 行/正文），
            // 分节符都在这里统一计数并重置状态，保证任何路径下跨节后标题区检测都会重启
            if (HasSectionBreak(p))
            {
                sectionIndex++;
                inCoverZone = false;
                if (sectionIndex > (hasCover ? 1 : 0))
                {
                    checkingTitle = true;
                    titleLinesFound = 0;
                    checkingDocNumber = false;
                    docNumberScanCount = 0;
                    checkingHeaderLines = false;
                    headerLinesFound = 0;
                }
            }
        }

        void 处理正文段落(Paragraph p)
        {
            var text = OpenXmlHelper.ParagraphText(p).Trim();

            if (inCoverZone && sectionIndex == 0)
            {
                if (!ShouldExitCoverZone(word.MainDocumentPart, p, text))
                {
                    return; // 分节符计数与状态重置由循环末尾的统一出口处理
                }

                inCoverZone = false;
            }

            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            // 正文区域的状态机：处理正文标题区和正文文号区
            if (checkingTitle && titleLinesFound < 3)
            {
                if (IsParagraphTitle(word.MainDocumentPart, p))
                {
                    titleLinesFound++;
                    FormatAsTitle(p);
                    return; // 必须 return，避免应用下方普通的 21 磅行距和首行缩进
                }
                else
                {
                    checkingTitle = false;
                    checkingDocNumber = true; // 标题区结束，开始检查文号
                    docNumberScanCount = 0;
                }
            }
            else if (checkingTitle)
            {
                checkingTitle = false;
                checkingDocNumber = true;
                docNumberScanCount = 0;
            }

            if (checkingDocNumber)
            {
                docNumberScanCount++;
                // 统一用归一化后的文本匹配，避免“川 华 信”夹空格漏判
                var normalizedDocNumberText = OpenXmlHelper.NormalizeText(text);
                var matchedByPrimaryRule = titleLinesFound > 0 && normalizedDocNumberText.Contains("川华信") && text.Length < 50 && !normalizedDocNumberText.Contains("事务所");
                var matchedByFallbackRule = !text.Contains("事务所") && text.Length < 60 && DocNumberRegex.IsMatch(normalizedDocNumberText);
                if (matchedByPrimaryRule || matchedByFallbackRule)
                {
                    FormatAsDocNumber(p);
                    checkingDocNumber = false;
                    docNumberScanCount = 0;
                    checkingHeaderLines = true;
                    headerLinesFound = 0;
                    return;
                }
                if (docNumberScanCount >= 5)
                {
                    checkingDocNumber = false;
                    docNumberScanCount = 0;
                }
            }

            if (checkingHeaderLines)
            {
                if (headerLinesFound < 2 && IsHeaderLineText(text))
                {
                    FormatAsHeaderLine(p);
                    headerLinesFound++;
                    if (headerLinesFound >= 2)
                    {
                        checkingHeaderLines = false;
                        headerLinesFound = 0;
                    }
                    return;
                }
                checkingHeaderLines = false;
                headerLinesFound = 0;
            }

            var headingLevel = ResolveHeadingLevel(word.MainDocumentPart, p, text);
            var hasAutomaticNumbering = headingLevel == 0 && OpenXmlHelper.段落存在编号(word.MainDocumentPart, p);

            if (headingLevel == 3)
            {
                var normalized = H3Separator.Replace(text, "$1．", 1);
                if (normalized != text)
                {
                    ReplaceParagraphVisibleText(p, normalized);
                    text = normalized;
                }
            }

            OpenXmlHelper.EnsureParagraphProperties(p);
            if (headingLevel is 1 or 2 or 3)
            {
                OpenXmlHelper.物化编号到段落(word.MainDocumentPart, p);
            }
            var pPr = p.ParagraphProperties!;
            
            if (headingLevel == 1)
            {
                ClearHeadingParagraphContaminants(pPr);
                // BeforeLines/AfterLines 存在时 Before/After 会被 Word 忽略，只保留 Lines 版本
                pPr.SpacingBetweenLines = new SpacingBetweenLines {
                    LineRule = LineSpacingRuleValues.Auto, Line = "360",
                    BeforeLines = 100, AfterLines = 50,
                    BeforeAutoSpacing = false, AfterAutoSpacing = false
                };
                pPr.Indentation = CreateHeadingIndentation(1);
                pPr.Justification = new Justification { Val = JustificationValues.Both };
                pPr.KeepNext = new KeepNext();
                pPr.OutlineLevel = new OutlineLevel { Val = 0 };
                ForceNumberingIndent(word.MainDocumentPart, p, 1, 长编号段落索引);
            }
            else if (headingLevel == 2)
            {
                ClearHeadingParagraphContaminants(pPr);
                pPr.SpacingBetweenLines = new SpacingBetweenLines {
                    LineRule = LineSpacingRuleValues.Auto, Line = "360",
                    BeforeLines = 50, AfterLines = 25,
                    BeforeAutoSpacing = false, AfterAutoSpacing = false
                };
                pPr.Indentation = CreateHeadingIndentation(2);
                pPr.Justification = new Justification { Val = JustificationValues.Both };
                pPr.KeepNext = new KeepNext();
                pPr.OutlineLevel = new OutlineLevel { Val = 1 };
                ForceNumberingIndent(word.MainDocumentPart, p, 2, 长编号段落索引);
            }
            else if (headingLevel == 3)
            {
                ClearHeadingParagraphContaminants(pPr);
                pPr.SpacingBetweenLines = new SpacingBetweenLines {
                    LineRule = LineSpacingRuleValues.Auto, Line = "360",
                    BeforeLines = 25, AfterLines = 0,
                    BeforeAutoSpacing = false, AfterAutoSpacing = false
                };
                pPr.Indentation = CreateHeadingIndentation(3);
                pPr.Justification = new Justification { Val = JustificationValues.Both };
                pPr.KeepNext = new KeepNext();
                pPr.OutlineLevel = new OutlineLevel { Val = 2 };
                ForceNumberingIndent(word.MainDocumentPart, p, 3, 长编号段落索引);
            }
            else
            {
                if (hasAutomaticNumbering)
                {
                    OpenXmlHelper.物化编号到段落(word.MainDocumentPart, p);
                }
                ClearNormalBodyParagraphContaminants(pPr, !hasAutomaticNumbering);
                pPr.SpacingBetweenLines = new SpacingBetweenLines {
                    LineRule = LineSpacingRuleValues.Exact, Line = "420",
                    BeforeLines = 0, AfterLines = 0,
                    BeforeAutoSpacing = false, AfterAutoSpacing = false
                };
                pPr.Indentation = CreateBodyIndentation();
                pPr.Justification = new Justification { Val = JustificationValues.Both };
                pPr.OutlineLevel = new OutlineLevel { Val = 9 };
            }

            foreach (var run in OpenXmlHelper.Runs(p).ToList())
            {
                run.RunProperties ??= new RunProperties();
                
                var scale = run.RunProperties.GetFirstChild<CharacterScale>();
                if (scale != null) scale.Remove();
                var fitText = run.RunProperties.GetFirstChild<FitText>();
                if (fitText != null) fitText.Remove();
                var spacing = run.RunProperties.GetFirstChild<Spacing>();
                if (spacing != null) spacing.Remove();
                var position = run.RunProperties.GetFirstChild<Position>();
                if (position != null) position.Remove();
                run.RunProperties.RunStyle?.Remove();

                run.RunProperties.FontSize = new FontSize { Val = "24" };
                run.RunProperties.FontSizeComplexScript = new FontSizeComplexScript { Val = "24" };
                OpenXmlHelper.SetRunFonts(run.RunProperties, "宋体", "Times New Roman");
                
                if (headingLevel is 1 or 2 or 3)
                {
                    run.RunProperties.Bold = new Bold { Val = true };
                    run.RunProperties.BoldComplexScript = new BoldComplexScript { Val = true };
                }
                else 
                {
                    run.RunProperties.Bold = new Bold { Val = false };
                    run.RunProperties.BoldComplexScript = new BoldComplexScript { Val = false };
                }
                
                run.RunProperties.Italic = new Italic { Val = false };
                run.RunProperties.ItalicComplexScript = new ItalicComplexScript { Val = false };
                var underline = run.RunProperties.GetFirstChild<Underline>();
                if (underline != null) underline.Remove();
            }

            if (headingLevel is 1 or 2 or 3)
            {
                var markRpr = pPr.ParagraphMarkRunProperties;
                if (markRpr == null)
                {
                    markRpr = new ParagraphMarkRunProperties();
                    pPr.ParagraphMarkRunProperties = markRpr;
                }
                markRpr.RemoveAllChildren();
                markRpr.Append(
                    new RunFonts
                    {
                        Ascii = "Times New Roman",
                        HighAnsi = "Times New Roman",
                        ComplexScript = "Times New Roman",
                        EastAsia = "宋体"
                    },
                    new Bold { Val = true },
                    new BoldComplexScript { Val = true },
                    new FontSize { Val = "24" },
                    new FontSizeComplexScript { Val = "24" });
            }

        }
    }

    private static bool HasSectionBreak(Paragraph p)
    {
        return p.ParagraphProperties?.GetFirstChild<SectionProperties>() is not null;
    }

    private static bool ShouldExitCoverZone(MainDocumentPart? mainPart, Paragraph paragraph, string text)
    {
        var normalized = OpenXmlHelper.NormalizeText(text);
        if (string.IsNullOrEmpty(normalized))
        {
            return false;
        }

        if (normalized.Contains("目录", StringComparison.Ordinal))
        {
            return true;
        }

        if (H1.IsMatch(normalized) || H2.IsMatch(normalized) || H3.IsMatch(normalized))
        {
            return true;
        }

        var visibleNumberingLevel = OpenXmlHelper.ResolveHeadingLevelByVisibleNumbering(mainPart, paragraph);
        return visibleNumberingLevel is 1 or 2 or 3;
    }

    private static void ClearNormalBodyParagraphContaminants(ParagraphProperties pPr, bool removeNumbering)
    {
        if (removeNumbering)
        {
            pPr.GetFirstChild<NumberingProperties>()?.Remove();
        }
        pPr.GetFirstChild<Tabs>()?.Remove();
        if (pPr.KeepNext != null) pPr.KeepNext.Remove();
        if (pPr.KeepLines != null) pPr.KeepLines.Remove();
        if (pPr.PageBreakBefore != null) pPr.PageBreakBefore.Remove();
        if (pPr.ContextualSpacing != null) pPr.ContextualSpacing.Remove();
        if (pPr.OutlineLevel != null) pPr.OutlineLevel.Remove();
    }

    private static void ClearHeadingParagraphContaminants(ParagraphProperties pPr)
    {
        pPr.GetFirstChild<Tabs>()?.Remove();
        if (pPr.KeepNext != null) pPr.KeepNext.Remove();
        if (pPr.KeepLines != null) pPr.KeepLines.Remove();
        if (pPr.PageBreakBefore != null) pPr.PageBreakBefore.Remove();
        if (pPr.ContextualSpacing != null) pPr.ContextualSpacing.Remove();
        if (pPr.OutlineLevel != null) pPr.OutlineLevel.Remove();
    }

    private static Indentation CreateHeadingIndentation(int headingLevel)
    {
        return new Indentation
        {
            FirstLine = headingLevel == 3 ? "480" : "0",
            FirstLineChars = headingLevel == 3 ? 200 : null,
            Left = "0",
            Right = "0",
            Start = "0",
            LeftChars = 0,
            RightChars = 0,
            StartCharacters = 0,
            Hanging = null,
            HangingChars = null
        };
    }

    private static Indentation CreateBodyIndentation()
    {
        return new Indentation
        {
            FirstLine = "480",
            FirstLineChars = 200,
            Left = "0",
            Right = "0",
            Start = "0",
            LeftChars = 0,
            RightChars = 0,
            StartCharacters = 0,
            Hanging = null,
            HangingChars = null
        };
    }

    private static bool IsParagraphTitle(MainDocumentPart? mainPart, Paragraph p)
    {
        var paragraphText = OpenXmlHelper.ParagraphText(p).Trim();
        if (ResolveHeadingLevel(mainPart, p, paragraphText, false) != 0) return false;

        var runs = p.Descendants<Run>().Where(r => !string.IsNullOrWhiteSpace(OpenXmlHelper.提取可见文本(r))).ToList();
        if (!runs.Any()) return false;

        var pPr = p.ParagraphProperties;
        var paragraphMarkRpr = pPr?.ParagraphMarkRunProperties;
        var paragraphMarkBold = IsBold(paragraphMarkRpr);
        var paragraphMarkSize = GetFontSize(paragraphMarkRpr);
        var paragraphMarkStyleId = paragraphMarkRpr?.GetFirstChild<RunStyle>()?.Val?.Value;
        var paragraphStyleId = pPr?.ParagraphStyleId?.Val?.Value;

        foreach (var run in runs)
        {
            var rPr = run.RunProperties;
            var runStyleId = rPr?.RunStyle?.Val?.Value;
            var isBold =
                IsBold(rPr)
                || (runStyleId != null && OpenXmlHelper.IsStyleBold(mainPart, StyleValues.Character, runStyleId))
                || paragraphMarkBold
                || (paragraphMarkStyleId != null && OpenXmlHelper.IsStyleBold(mainPart, StyleValues.Character, paragraphMarkStyleId))
                || (paragraphStyleId != null && OpenXmlHelper.IsStyleBold(mainPart, StyleValues.Paragraph, paragraphStyleId));
            var size =
                GetFontSize(rPr)
                ?? OpenXmlHelper.GetStyleFontSize(mainPart, StyleValues.Character, runStyleId)
                ?? paragraphMarkSize
                ?? OpenXmlHelper.GetStyleFontSize(mainPart, StyleValues.Character, paragraphMarkStyleId)
                ?? OpenXmlHelper.GetStyleFontSize(mainPart, StyleValues.Paragraph, paragraphStyleId);
            var isLarge = size.HasValue && size.Value > 24;

            if (isBold && isLarge) return true;
        }

        return false;
    }

    private static int ResolveHeadingLevel(MainDocumentPart? mainPart, Paragraph p, string text, bool allowStyleAndOutline = true)
    {
        var headingLevel = 0;
        if (text.Length <= 30)
        {
            if (H1.IsMatch(text)) return 1;
            if (H2.IsMatch(text)) return 2;
            if (H3.IsMatch(text)) return 3;
        }

        headingLevel = OpenXmlHelper.ResolveHeadingLevelByVisibleNumbering(mainPart, p);
        if (headingLevel != 0) return headingLevel;

        if (!allowStyleAndOutline || text.Length > 30) return 0;

        var outlineVal = OpenXmlHelper.ResolveOutlineLevel(mainPart, p);
        if (outlineVal == 0) return 1;
        if (outlineVal == 1) return 2;
        if (outlineVal == 2) return 3;

        var styleVal = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        return OpenXmlHelper.ResolveHeadingLevelByStyleId(styleVal);
    }

    private static bool IsBold(OpenXmlElement? rPr)
    {
        if (rPr == null) return false;
        var bold = rPr.GetFirstChild<Bold>();
        return bold != null && (bold.Val == null || bold.Val.Value);
    }

    private static int? GetFontSize(OpenXmlElement? rPr)
    {
        var sz = rPr?.GetFirstChild<FontSize>()?.Val?.Value;
        if (string.IsNullOrWhiteSpace(sz)) return null;
        return int.TryParse(sz, out var value) ? value : null;
    }

    private static void ReplaceParagraphVisibleText(Paragraph paragraph, string text)
    {
        // 段落里若含书签/批注锚点等 annotation 元素，整体清空会把它们一并删掉，保守起见跳过替换
        if (paragraph.Descendants<BookmarkStart>().Any()
            || paragraph.Descendants<BookmarkEnd>().Any()
            || paragraph.Descendants<CommentRangeStart>().Any()
            || paragraph.Descendants<CommentRangeEnd>().Any())
        {
            return;
        }

        // 保留段落属性，避免清除后在后续 EnsureParagraphProperties 处丢失已有格式
        var pPr = paragraph.ParagraphProperties;

        // 彻底清除段落内所有子元素（包括嵌套在 hyperlink/sdt/ins 等结构里的 Run）
        paragraph.RemoveAllChildren();

        // 恢复段落属性（如果原本有的话）
        if (pPr != null)
        {
            paragraph.Append(pPr);
        }

        paragraph.Append(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static bool IsHeaderLineText(string text)
    {
        var normalized = OpenXmlHelper.NormalizeText(text);
        if (string.IsNullOrWhiteSpace(normalized)) return false;
        if (!normalized.EndsWith('：') && !normalized.EndsWith(':')) return false;
        if (normalized.Length > 60) return false;
        if (normalized.Length < 4) return false;
        if (normalized.Contains("中国注册会计师") || normalized.Contains("四川华信")) return false;
        if (normalized.StartsWith("（") || normalized.StartsWith("(")) return false;
        return HeaderOrgKeywords.Any(kw => normalized.Contains(kw) && normalized.Length - kw.Length <= 20);
    }

    private static void FormatAsTitle(Paragraph p)
    {
        var pPr = ResetTitleParagraphProperties(p);

        pPr.SpacingBetweenLines = new SpacingBetweenLines
        {
            Before = "0",
            After = "0",
            BeforeLines = 50,
            AfterLines = 50,
            LineRule = LineSpacingRuleValues.Auto,
            Line = "360",
            BeforeAutoSpacing = false,
            AfterAutoSpacing = false
        };

        pPr.Justification = new Justification { Val = JustificationValues.Center };
        pPr.Indentation = new Indentation { FirstLine = "0", Left = "0", Right = "0", Start = "0", LeftChars = 0, RightChars = 0, StartCharacters = 0, FirstLineChars = 0 };

        foreach (var run in OpenXmlHelper.Runs(p))
        {
            ApplyTitleRunProperties(run, "32", true);
        }
    }

    private static void FormatAsDocNumber(Paragraph p)
    {
        var pPr = ResetTitleParagraphProperties(p);

        pPr.SpacingBetweenLines = new SpacingBetweenLines
        {
            Before = "0",
            After = "0",
            BeforeLines = 25,
            AfterLines = 25,
            LineRule = LineSpacingRuleValues.Auto,
            Line = "360",
            BeforeAutoSpacing = false,
            AfterAutoSpacing = false
        };

        pPr.Justification = new Justification { Val = JustificationValues.Right };
        pPr.Indentation = new Indentation { FirstLine = "0", Left = "0", Right = "0", Start = "0", LeftChars = 0, RightChars = 0, StartCharacters = 0, FirstLineChars = 0 };

        foreach (var run in OpenXmlHelper.Runs(p))
        {
            ApplyTitleRunProperties(run, "24", false);
        }
    }

    private static void FormatAsHeaderLine(Paragraph p)
    {
        var pPr = ResetTitleParagraphProperties(p);

        pPr.SpacingBetweenLines = new SpacingBetweenLines
        {
            BeforeLines = 0,
            AfterLines = 0,
            LineRule = LineSpacingRuleValues.Auto,
            Line = "360",
            Before = "0",
            After = "0",
            BeforeAutoSpacing = false,
            AfterAutoSpacing = false
        };

        pPr.Justification = new Justification { Val = JustificationValues.Left };
        pPr.Indentation = new Indentation { FirstLine = "0", Left = "0", Right = "0", Start = "0", LeftChars = 0, RightChars = 0, StartCharacters = 0, FirstLineChars = 0 };

        foreach (var run in OpenXmlHelper.Runs(p))
        {
            ApplyTitleRunProperties(run, "24", true);
        }
    }

    private static ParagraphProperties ResetTitleParagraphProperties(Paragraph paragraph)
    {
        var sectionProperties = paragraph.ParagraphProperties?.GetFirstChild<SectionProperties>()?.CloneNode(true) as SectionProperties;
        var numberingProperties = paragraph.ParagraphProperties?.GetFirstChild<NumberingProperties>()?.CloneNode(true) as NumberingProperties;
        var paragraphStyleId = paragraph.ParagraphProperties?.GetFirstChild<ParagraphStyleId>()?.CloneNode(true) as ParagraphStyleId;
        var pPr = new ParagraphProperties();
        if (paragraphStyleId != null)
        {
            pPr.Append(paragraphStyleId);
        }
        if (numberingProperties != null)
        {
            pPr.Append(numberingProperties);
        }
        if (sectionProperties != null)
        {
            pPr.Append(sectionProperties);
        }
        paragraph.ParagraphProperties = pPr;
        return pPr;
    }

    /// <summary>
    /// 在现有 rPr 上按字段覆盖字体/字号/加粗等，而不是整体替换 RunProperties，
    /// 避免丢掉 Vanish/WebHidden（隐藏文字）、VertAlign（上下标，如脚注引用符）、rStyle 等既有设置
    /// </summary>
    private static void ApplyTitleRunProperties(Run run, string fontSize, bool bold)
    {
        run.RunProperties ??= new RunProperties();
        var rPr = run.RunProperties;

        rPr.FontSize = new FontSize { Val = fontSize };
        rPr.FontSizeComplexScript = new FontSizeComplexScript { Val = fontSize };
        rPr.Bold = new Bold { Val = bold };
        rPr.BoldComplexScript = new BoldComplexScript { Val = bold };
        rPr.Italic = new Italic { Val = false };
        rPr.ItalicComplexScript = new ItalicComplexScript { Val = false };
        OpenXmlHelper.SetRunFonts(rPr, "宋体", "Times New Roman");
    }

    private static void ForceNumberingIndent(MainDocumentPart? mainPart, Paragraph p, int headingLevel,
        IReadOnlyDictionary<(int AbstractNumId, int Ilvl), List<string>> 长编号段落索引)
    {
        if (mainPart == null) return;
        var numPr = OpenXmlHelper.ResolveParagraphNumberingProperties(mainPart, p);
        if (numPr == null) return;

        var numIdVal = numPr.NumberingId?.Val?.Value;
        var ilvlVal = numPr.NumberingLevelReference?.Val?.Value;
        if (numIdVal == null || ilvlVal == null) return;

        var numberingPart = mainPart.NumberingDefinitionsPart;
        if (numberingPart == null || numberingPart.Numbering == null) return;

        var numInstance = numberingPart.Numbering.Elements<NumberingInstance>().FirstOrDefault(n => n.NumberID?.Value == numIdVal);
        var abstractNumId = numInstance?.AbstractNumId?.Val?.Value;
        if (abstractNumId == null) return;

        var abstractNum = numberingPart.Numbering.Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId);
        if (abstractNum == null) return;

        var lvl = abstractNum.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == ilvlVal);
        if (lvl == null) return;
        // 同一 abstractNumId（含绑定同一 abstractNum 的其他 numId）下若已有长文本段落共享该编号，
        // 说明这是正文长段落共用的编号，不能动它的缩进和编号字体
        if (存在共享长编号段落(长编号段落索引, abstractNumId.Value, ilvlVal.Value, headingLevel)) return;

        var lvlRpr = lvl.NumberingSymbolRunProperties;
        if (lvlRpr == null)
        {
            lvlRpr = new NumberingSymbolRunProperties();
            lvl.NumberingSymbolRunProperties = lvlRpr;
        }
        lvlRpr.RunFonts = new RunFonts
        {
            Ascii = "Times New Roman",
            HighAnsi = "Times New Roman",
            ComplexScript = "Times New Roman",
            EastAsia = "宋体"
        };
        lvlRpr.Bold = new Bold();
        lvlRpr.FontSize = new FontSize { Val = "24" };
        lvlRpr.FontSizeComplexScript = new FontSizeComplexScript { Val = "24" };

        var pPr = lvl.PreviousParagraphProperties;
        if (pPr == null)
        {
            pPr = new PreviousParagraphProperties();
            lvl.PreviousParagraphProperties = pPr;
        }

        if (pPr.Indentation == null)
        {
            pPr.Indentation = new Indentation();
        }

        // 修改编号后的后缀字符，默认是 Tab，会导致文本看起来被缩进
        var suffix = lvl.Elements<LevelSuffix>().FirstOrDefault();
        if (suffix == null)
        {
            suffix = new LevelSuffix();
            // w:lvl 子元素顺序：start→numFmt→lvlRestart→pStyle→isLgl→suff→lvlText→…，suff 必须位于 lvlText 之前
            if (lvl.LevelText != null)
            {
                lvl.InsertBefore(suffix, lvl.LevelText);
            }
            else
            {
                // 无 lvlText 时按粒子顺序找 suff 之前最后一个已有元素作为锚点
                OpenXmlElement? after = lvl.IsLegalNumberingStyle;
                after ??= lvl.ParagraphStyleIdInLevel;
                after ??= lvl.LevelRestart;
                after ??= lvl.NumberingFormat;
                after ??= lvl.StartNumberingValue;
                if (after != null) lvl.InsertAfter(suffix, after);
                else lvl.InsertAt(suffix, 0); // 没有任何前置锚点时，suff 放最前（后续元素均在 suff 之后，顺序合法）
            }
        }
        suffix.Val = LevelSuffixValues.Nothing; // 或者使用 Space

        if (headingLevel == 1 || headingLevel == 2)
        {
            pPr.Indentation.Left = "0";
            pPr.Indentation.Start = "0";
            pPr.Indentation.FirstLine = "0";
            pPr.Indentation.Hanging = null;
            pPr.Indentation.LeftChars = null;
            pPr.Indentation.StartCharacters = null;
            pPr.Indentation.FirstLineChars = null;
            pPr.Indentation.HangingChars = null;
        }
        else if (headingLevel == 3)
        {
            pPr.Indentation.Left = "0";
            pPr.Indentation.Start = "0";
            pPr.Indentation.FirstLine = "480";
            pPr.Indentation.Hanging = null;
            pPr.Indentation.LeftChars = null;
            pPr.Indentation.StartCharacters = null;
            pPr.Indentation.FirstLineChars = 200;
            pPr.Indentation.HangingChars = null;
        }
    }

    /// <summary>
    /// 循环前一次性扫描全文，建立 (abstractNumId, ilvl) → 长文本段落归一化文本列表 的索引。
    /// 只收录归一化后超过 30 字的段落文本（标题段落文本必然 ≤30 字，不会误入索引）。
    /// 按 abstractNumId 建键，绑定同一 abstractNum 的不同 numId 段落也在保护范围内。
    /// </summary>
    private static Dictionary<(int AbstractNumId, int Ilvl), List<string>> 建立长编号段落索引(MainDocumentPart? mainPart)
    {
        var index = new Dictionary<(int AbstractNumId, int Ilvl), List<string>>();
        var body = mainPart?.Document?.Body;
        var numbering = mainPart?.NumberingDefinitionsPart?.Numbering;
        if (body == null || numbering == null) return index;

        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            var numPr = OpenXmlHelper.ResolveParagraphNumberingProperties(mainPart, paragraph);
            var numId = numPr?.NumberingId?.Val?.Value;
            var ilvl = numPr?.NumberingLevelReference?.Val?.Value;
            if (numId == null || ilvl == null) continue;

            var abstractNumId = numbering.Elements<NumberingInstance>()
                .FirstOrDefault(n => n.NumberID?.Value == numId.Value)?.AbstractNumId?.Val?.Value;
            if (abstractNumId == null) continue;

            var normalizedText = OpenXmlHelper.NormalizeText(OpenXmlHelper.ParagraphText(paragraph));
            if (normalizedText.Length <= 30) continue;

            var key = (abstractNumId.Value, ilvl.Value);
            if (!index.TryGetValue(key, out var list))
            {
                list = new List<string>();
                index[key] = list;
            }
            list.Add(normalizedText);
        }

        return index;
    }

    private static bool 存在共享长编号段落(
        IReadOnlyDictionary<(int AbstractNumId, int Ilvl), List<string>> 长编号段落索引,
        int abstractNumId, int ilvl, int headingLevel)
    {
        if (!长编号段落索引.TryGetValue((abstractNumId, ilvl), out var 长文本列表)) return false;

        return headingLevel switch
        {
            1 => 长文本列表.Any(H1.IsMatch),
            2 => 长文本列表.Any(H2.IsMatch),
            3 => 长文本列表.Any(H3.IsMatch),
            _ => false
        };
    }
}
