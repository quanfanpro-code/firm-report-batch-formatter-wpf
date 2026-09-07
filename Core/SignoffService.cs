using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace FirmFormatter.OpenXml.Core;

public sealed class SignoffService
{
    private static readonly TimeSpan 正则超时 = TimeSpan.FromSeconds(1);
    private static readonly Regex CpaRegex = new(@"中国注册\s*会计师.*", RegexOptions.Compiled, 正则超时);
    private static readonly Regex TemplateDateRegex = new(@"^(?:20\d{2}|[〇零一二三四五六七八九十]{2,4})年(?:[〇零一二三四五六七八九十\dXx×]{1,3})月(?:[〇零一二三四五六七八九十\dXx×]{1,3})日$", RegexOptions.Compiled, 正则超时);
    private readonly FirmRuleProfile _ruleProfile = FirmRuleProfile.Default;
    private readonly string _firmStartBoundary;
    private readonly string _cityBoundary;

    public SignoffService()
    {
        var template = FirmTemplateProfile.Default;
        _firmStartBoundary = NormalizeBoundaryText(template.事务所落款头部);
        _cityBoundary = NormalizeBoundaryText(template.落款城市行);
    }

    public List<Paragraph> IdentifySignoffParagraphs(WordprocessingDocument word, bool hasCover)
    {
        // 防卫：任一边界关键词配置为空时，string.Contains("", Ordinal) 恒真，
        // 第一个非空段落就会被误判为落款起点，此时直接放弃识别
        if (_firmStartBoundary.Length == 0 || _cityBoundary.Length == 0) return [];

        var body = word.MainDocumentPart?.Document?.Body;
        if (body is null) return [];

        var allProbes = body.Descendants<Paragraph>()
            .Select((paragraph, index) => new ParagraphProbe(
                index,
                paragraph,
                VisibleParagraphText(paragraph).Trim()))
            .ToList();

        // 封面区 = 第一个分节（首个分节符段落及其之前）；
        // hasCover 时跳过封面区段落，避免封面上的事务所名/日期行参与落款误匹配
        var coverEndIndex = hasCover
            ? allProbes.FirstOrDefault(x =>
                x.Paragraph.ParagraphProperties?.GetFirstChild<SectionProperties>() is not null)?.Index ?? -1
            : -1;

        var paragraphs = allProbes
            .Where(x => !string.IsNullOrWhiteSpace(x.Text))
            .Where(x => x.Index > coverEndIndex)
            .ToList();
        if (paragraphs.Count == 0) return [];

        SignoffRange? matchedRange = null;
        foreach (var startProbe in paragraphs.Where(x =>
                     x.NormalizedText.Contains(_firmStartBoundary, StringComparison.Ordinal)))
        {
            var cityProbe = paragraphs.FirstOrDefault(x =>
                x.Index >= startProbe.Index &&
                x.NormalizedText.Contains(_cityBoundary, StringComparison.Ordinal));
            if (cityProbe is null) continue;

            var dateProbe = paragraphs.FirstOrDefault(x =>
                x.Index >= cityProbe.Index &&
                IsTemplateDateLine(x.NormalizedText));
            if (dateProbe is null) continue;

            var cpaLineCount = paragraphs.Count(x =>
                x.Index >= startProbe.Index &&
                x.Index <= dateProbe.Index &&
                CpaRegex.IsMatch(x.Text));
            if (cpaLineCount < 2) continue;

            matchedRange = new SignoffRange(startProbe.Index, dateProbe.Index);
        }

        if (matchedRange is null) return [];

        return paragraphs
            .Where(x => x.Index >= matchedRange.StartIndex && x.Index <= matchedRange.EndIndex)
            .Select(x => x.Paragraph)
            .ToList();
    }

    public void Apply(WordprocessingDocument word, List<Paragraph> signoffParagraphs, CancellationToken cancellationToken = default)
    {
        if (signoffParagraphs == null || signoffParagraphs.Count == 0) return;

        // 契约：末段按日期行处理。传入列表不保证有序（变量原名 sorted 实未排序），
        // 这里统一按文档顺序重排，确保 dateIndex 落在真正的最后一段
        var sortedParas = 按文档顺序重排(word, signoffParagraphs);
        var dateIndex = sortedParas.Count - 1;
        var steppedFirstLines = new[] { "0", "480", "960" };

        for (var i = 0; i < sortedParas.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var p = sortedParas[i];
            OpenXmlHelper.EnsureParagraphProperties(p);
            var pPr = p.ParagraphProperties!;

            // 1. 强制落款区行距为 42 磅 (840 twips)，并彻底清空所有段前段后间距
            pPr.SpacingBetweenLines = new SpacingBetweenLines
            {
                LineRule = LineSpacingRuleValues.Exact,
                Line = _ruleProfile.落款行距Twips,
                Before = "0",
                BeforeLines = 0,
                After = "0",
                AfterLines = 0
            };

            // 2. 彻底清除不需要的属性（防守加固）
            if (pPr.KeepNext != null) pPr.KeepNext.Remove();
            if (pPr.OutlineLevel != null) pPr.OutlineLevel.Remove();

            // 统一落款区字体为中文宋体，英文Times New Roman，并清除杂项污染
            foreach (var run in OpenXmlHelper.Runs(p))
            {
                if (run.Descendants<FootnoteReference>().Any() || run.Descendants<EndnoteReference>().Any())
                {
                    continue;
                }

                run.RunProperties ??= new RunProperties();
                OpenXmlHelper.SetRunFonts(run.RunProperties, _ruleProfile.中文字体, _ruleProfile.西文字体);
                run.RunProperties.Bold = new Bold { Val = false };
                run.RunProperties.BoldComplexScript = new BoldComplexScript { Val = false };
                run.RunProperties.Italic = new Italic { Val = false };
                run.RunProperties.ItalicComplexScript = new ItalicComplexScript { Val = false };
                run.RunProperties.FontSize = new FontSize { Val = _ruleProfile.正文字号HalfPoint };
                run.RunProperties.FontSizeComplexScript = new FontSizeComplexScript { Val = _ruleProfile.正文字号HalfPoint };

                var scale = run.RunProperties.GetFirstChild<CharacterScale>();
                if (scale != null) scale.Remove();
                var fitText = run.RunProperties.GetFirstChild<FitText>();
                if (fitText != null) fitText.Remove();
                var spacing = run.RunProperties.GetFirstChild<Spacing>();
                if (spacing != null) spacing.Remove();
                var position = run.RunProperties.GetFirstChild<Position>();
                if (position != null) position.Remove();
                run.RunProperties.RunStyle?.Remove();
            }

            var text = VisibleParagraphText(p);

            // 如果是最后一行日期，保持右对齐
            if (i == dateIndex)
            {
                pPr.Tabs = null;
                pPr.Indentation = CreateSignoffIndentation("0");
                pPr.Justification = new Justification { Val = JustificationValues.Right };
            }
            else
            {
                ApplySteppedIndentation(pPr, i < steppedFirstLines.Length ? steppedFirstLines[i] : "0");
                pPr.Justification = new Justification { Val = JustificationValues.Left };

                if (CpaRegex.IsMatch(text))
                {
                    EnsureRightAlignedSignatureTab(pPr);
                    NormalizeCpaSignatureLine(p);
                }
                else
                {
                    pPr.Tabs = null;
                }
            }
        }
    }

    private static List<Paragraph> 按文档顺序重排(WordprocessingDocument word, List<Paragraph> paragraphs)
    {
        var body = word.MainDocumentPart?.Document?.Body;
        if (body is null) return paragraphs.ToList();

        var order = new Dictionary<Paragraph, int>();
        var index = 0;
        foreach (var p in body.Descendants<Paragraph>())
        {
            order[p] = index++;
        }

        // 不在正文中的段落排到最后，不会被误当作首段/末段
        return paragraphs
            .OrderBy(p => order.TryGetValue(p, out var i) ? i : int.MaxValue)
            .ToList();
    }

    private void NormalizeCpaSignatureLine(Paragraph paragraph)
    {
        // 含链接、域、修订、换行、隐藏文字、脚注等结构时只做安全的格式覆盖，不清空重建内容。
        if (paragraph.Descendants().Any(element =>
                element is Hyperlink or SimpleField or FieldCode or Drawing or InsertedRun or DeletedRun
                    or Break or TabChar or Vanish or WebHidden or FootnoteReference or EndnoteReference
                || element.LocalName is "sdt" or "oMath" or "oMathPara" or "object" or "pict"
                    or "commentRangeStart" or "commentRangeEnd" or "commentReference"))
        {
            return;
        }

        var fullText = VisibleParagraphText(paragraph);
        var match = CpaRegex.Match(fullText);
        if (!match.Success) return;

        var right = CleanRightPart(match.Value);
        if (string.IsNullOrWhiteSpace(right)) return;

        var cleanProps = CreateCleanSignatureRunProperties();
        var left = fullText[..match.Index].TrimEnd(' ', '\t', '\u3000');

        // 保留 BookmarkStart/BookmarkEnd，避免破坏书签锚点，其余内容元素清掉重建
        foreach (var child in paragraph.ChildElements
                     .Where(child => child is not ParagraphProperties
                                     and not BookmarkStart
                                     and not BookmarkEnd)
                     .ToList())
        {
            child.Remove();
        }

        if (!string.IsNullOrWhiteSpace(left))
        {
            var leftRun = new Run();
            leftRun.RunProperties = (RunProperties)cleanProps.CloneNode(true);
            leftRun.AppendChild(new Text(left) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.AppendChild(leftRun);
        }

        var rightRun = new Run();
        rightRun.RunProperties = (RunProperties)cleanProps.CloneNode(true);
        rightRun.AppendChild(new TabChar());
        rightRun.AppendChild(new Text(right) { Space = SpaceProcessingModeValues.Preserve });
        paragraph.AppendChild(rightRun);
    }

    private void EnsureRightAlignedSignatureTab(ParagraphProperties paragraphProperties)
    {
        paragraphProperties.Tabs = new Tabs(new TabStop
        {
            Val = TabStopValues.Right,
            Position = _ruleProfile.注册会计师制表位Twips
        });
    }

    private static void ApplySteppedIndentation(ParagraphProperties paragraphProperties, string firstLine)
    {
        // 缇值与字符单位必须成对写：Word 规则是 FirstLineChars 优先于 FirstLine，
        // 只写缇值时样式链上残留的字符单位缩进（如 Normal 的 200）会架空缇值
        paragraphProperties.Indentation = CreateSignoffIndentation(firstLine);
    }

    private static Indentation CreateSignoffIndentation(string firstLine)
    {
        return new Indentation
        {
            Left = "0",
            Right = "0",
            Start = "0",
            LeftChars = 0,
            RightChars = 0,
            StartCharacters = 0,
            FirstLine = firstLine,
            FirstLineChars = int.Parse(firstLine) / 480 * 200,
            Hanging = null,
            HangingChars = null
        };
    }

    private RunProperties CreateCleanSignatureRunProperties()
    {
        var runProperties = new RunProperties
        {
            Bold = new Bold { Val = false },
            BoldComplexScript = new BoldComplexScript { Val = false },
            Italic = new Italic { Val = false },
            ItalicComplexScript = new ItalicComplexScript { Val = false },
            FontSize = new FontSize { Val = _ruleProfile.正文字号HalfPoint },
            FontSizeComplexScript = new FontSizeComplexScript { Val = _ruleProfile.正文字号HalfPoint }
        };

        OpenXmlHelper.SetRunFonts(runProperties, _ruleProfile.中文字体, _ruleProfile.西文字体);
        return runProperties;
    }

    private static string CleanRightPart(string value)
    {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        return Regex.Replace(value, @"^[\s\u3000·]+", "");
    }

    private static bool IsTemplateDateLine(string normalizedText)
    {
        if (string.IsNullOrWhiteSpace(normalizedText)) return false;
        return TemplateDateRegex.IsMatch(normalizedText);
    }

    private static string VisibleParagraphText(Paragraph paragraph)
    {
        return OpenXmlHelper.提取可见文本(paragraph);
    }

    private static string NormalizeBoundaryText(string text)
    {
        var normalized = OpenXmlHelper.NormalizeText(text);
        if (normalized.Length == 0) return normalized;

        var chars = normalized.Select(NormalizeBoundaryChar).ToArray();
        return new string(chars);
    }

    private static char NormalizeBoundaryChar(char ch)
    {
        return ch switch
        {
            '（' => '(',
            '）' => ')',
            '：' => ':',
            '·' => '.',
            '•' => '.',
            '‧' => '.',
            '・' => '.',
            '．' => '.',
            '。' => '.',
            _ => ch
        };
    }

    private sealed record ParagraphProbe(int Index, Paragraph Paragraph, string Text)
    {
        public string NormalizedText { get; } = NormalizeBoundaryText(Text);
    }

    private sealed record SignoffRange(int StartIndex, int EndIndex);
}
