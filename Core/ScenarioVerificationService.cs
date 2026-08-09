using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FirmFormatter.OpenXml.Contracts;

namespace FirmFormatter.OpenXml.Core;

public sealed class ScenarioVerificationService
{
    private static readonly Regex 一级标题文本 = new(@"^[一二三四五六七八九十]+、", RegexOptions.Compiled);
    private static readonly Regex 二级标题文本 = new(@"^[（(][一二三四五六七八九十]+[)）]", RegexOptions.Compiled);
    private static readonly Regex 三级标题文本 = new(@"^\d{1,3}．", RegexOptions.Compiled);

    public static bool 支持场景(string scenarioName)
    {
        return NormalizeScenarioName(scenarioName) is
            "标题矩阵" or
            "正文编号矩阵" or
            "表格复杂文本矩阵" or
            "封面矩阵" or
            "纯封面矩阵" or
            "落款矩阵" or
            "共享编号模板防误伤矩阵" or
            "页眉页脚分节矩阵" or
            "文本框坑点矩阵" or
            "脚注尾注坑点矩阵" or
            "编号重启坑点矩阵" or
            "合并单元格坑点矩阵" or
            "落款复杂结构坑点矩阵" or
            "真实样本";
    }

    public ScenarioVerificationReportContract Verify(string docxPath, FirmDocumentContext context)
    {
        using var word = WordprocessingDocument.Open(docxPath, false);
        return Verify(word, context);
    }

    public ScenarioVerificationReportContract Verify(WordprocessingDocument word, FirmDocumentContext context)
    {
        var report = new ScenarioVerificationReportContract
        {
            Scenario = context.ScenarioName
        };

        report.Facts["验证封面判定"] = context.HasCover ? "1" : "0";
        report.Facts["验证纯封面判定"] = context.IsPureCoverDocument ? "1" : "0";

        var scenarioKey = NormalizeScenarioName(context.ScenarioName);
        switch (scenarioKey)
        {
            case "标题矩阵":
                VerifyHeadingMatrix(word, report);
                break;
            case "正文编号矩阵":
                VerifyBodyNumberingMatrix(word, report);
                break;
            case "表格复杂文本矩阵":
                VerifyTableMatrix(word, report);
                break;
            case "封面矩阵":
                VerifyCoverMatrix(word, report);
                break;
            case "纯封面矩阵":
                VerifyPureCoverMatrix(word, report, FirmRuleProfile.Default);
                break;
            case "落款矩阵":
                VerifySignoffMatrix(word, report, context);
                break;
            case "共享编号模板防误伤矩阵":
                VerifySharedNumberingMatrix(word, report);
                break;
            case "页眉页脚分节矩阵":
                VerifyHeaderFooterMatrix(word, report);
                break;
            case "文本框坑点矩阵":
                VerifyTextBoxPitfallMatrix(word, report, context);
                break;
            case "脚注尾注坑点矩阵":
                VerifyFootnoteEndnotePitfallMatrix(word, report, context);
                break;
            case "编号重启坑点矩阵":
                VerifyNumberingRestartPitfallMatrix(word, report, context);
                break;
            case "合并单元格坑点矩阵":
                VerifyMergedCellPitfallMatrix(word, report, context);
                break;
            case "落款复杂结构坑点矩阵":
                VerifyComplexSignoffPitfallMatrix(word, report, context);
                break;
            case "真实样本":
                VerifyRealSample(word, report);
                break;
            default:
                report.Issues.Add($"未识别的场景：{scenarioKey}");
                break;
        }

        return report;
    }

    private static void VerifyHeadingMatrix(WordprocessingDocument word, ScenarioVerificationReportContract report)
    {
        AssertHeading(word, report, "一、文本前缀一级标题", 1, true);
        AssertHeading(word, report, "（一）文本前缀二级标题", 2, true);
        AssertHeading(word, report, "1．文本前缀三级标题", 3, true);
        AssertHeading(word, report, "样式链一级标题", 1, true);
        AssertHeading(word, report, "样式链二级标题", 2, true);
        AssertHeading(word, report, "样式链三级标题", 3, true);
        AssertHeading(word, report, "段落自身一级标题", 1, true);
        AssertHeading(word, report, "段落自身二级标题", 2, true);
        AssertHeading(word, report, "段落自身三级标题", 3, true);
        AssertHeading(word, report, "仅大纲一级标题", 1, false);
        AssertHeading(word, report, "仅大纲二级标题", 2, false);
        AssertHeading(word, report, "仅大纲三级标题", 3, false);
    }

    private static void VerifyBodyNumberingMatrix(WordprocessingDocument word, ScenarioVerificationReportContract report)
    {
        AssertBodyNumberingPreserved(word, report, "这是正文式四级（直接编号）");
        AssertBodyNumberingPreserved(word, report, "这是正文式四级（样式链）");
        AssertBodyNumberingPreserved(word, report, "这是普通说明列表（直接编号）");
        AssertBodyNumberingPreserved(word, report, "这是普通说明列表（样式链）");
    }

    private static void VerifyTableMatrix(WordprocessingDocument word, ScenarioVerificationReportContract report)
    {
        var body = word.MainDocumentPart?.Document?.Body;
        var table = body?.Elements<Table>().FirstOrDefault();
        if (table == null)
        {
            report.Issues.Add("表格矩阵输出缺少表格");
            return;
        }

        var rows = table.Elements<TableRow>().ToList();
        if (rows.Count < 5)
        {
            report.Issues.Add("表格矩阵输出行数不足");
            return;
        }

        var row2Cells = rows[1].Elements<TableCell>().ToList();
        var row3Cells = rows[2].Elements<TableCell>().ToList();
        var row4Cells = rows[3].Elements<TableCell>().ToList();
        var row5Cells = rows[4].Elements<TableCell>().ToList();

        if (row2Cells.Count < 2 || row3Cells.Count < 2 || row4Cells.Count < 2 || row5Cells.Count < 2)
        {
            report.Issues.Add("表格矩阵输出列数不足");
            return;
        }

        AssertEquals(report, "表格序号列00123", "00123", NormalizeText(GetVisibleText(row2Cells[0])));
        AssertEquals(report, "表格数值1234", "1,234.00", NormalizeText(GetVisibleText(row2Cells[1])));
        AssertEquals(report, "表格拆分运行块数字", "1,234.00", NormalizeText(GetVisibleText(row3Cells[1])));

        var firstColJustification = row2Cells[0].Elements<Paragraph>().FirstOrDefault()?.ParagraphProperties?.Justification?.Val?.Value;
        if (firstColJustification != JustificationValues.Left)
        {
            report.Issues.Add("表格序号列没有按首列规则左对齐");
        }

        if (row4Cells[1].Elements<Paragraph>().Count() < 2)
        {
            report.Issues.Add("复杂单元格被粗暴重写，段落结构丢失");
        }

        AssertEquals(report, "表格序号列零值", "0", NormalizeText(GetVisibleText(row5Cells[0])));
        AssertEquals(report, "百分比格式", "12.50%", NormalizeText(GetVisibleText(row5Cells[1])));
    }

    private static void VerifyCoverMatrix(WordprocessingDocument word, ScenarioVerificationReportContract report)
    {
        var sections = OpenXmlHelper.收集分节(word.MainDocumentPart?.Document?.Body);
        report.Facts["分节数"] = sections.Count.ToString();
        if (sections.Count < 2)
        {
            report.Issues.Add("封面矩阵至少应包含封面节和正文节");
            return;
        }

        if (sections[0].Elements<HeaderReference>().Any() || sections[0].Elements<FooterReference>().Any())
        {
            report.Issues.Add("封面节页眉页脚未被清理");
        }
    }

    private static void VerifyPureCoverMatrix(WordprocessingDocument word, ScenarioVerificationReportContract report, FirmRuleProfile ruleProfile)
    {
        var sections = OpenXmlHelper.收集分节(word.MainDocumentPart?.Document?.Body);
        if (sections.Count == 0)
        {
            report.Issues.Add("纯封面矩阵缺少分节");
            return;
        }

        foreach (var section in sections)
        {
            var pageMargin = section.GetFirstChild<PageMargin>();
            if (pageMargin?.Bottom?.Value != (uint)ruleProfile.纯封面下边距Twips)
            {
                report.Issues.Add("纯封面矩阵下边距不是 2.2cm");
            }
        }
    }

    private static void VerifySignoffMatrix(WordprocessingDocument word, ScenarioVerificationReportContract report, FirmDocumentContext context)
    {
        var signoffParagraphs = new SignoffService()
            .IdentifySignoffParagraphs(word, context.HasCover);
        if (signoffParagraphs.Count == 0)
        {
            report.Issues.Add("落款矩阵未识别出落款区");
            return;
        }

        var cpaParagraphs = signoffParagraphs.Where(p => NormalizeText(GetVisibleText(p)).Contains("中国注册会计师", StringComparison.Ordinal)).ToList();
        if (cpaParagraphs.Count < 2)
        {
            report.Issues.Add("落款矩阵注册会计师行少于 2 行");
        }

        foreach (var paragraph in cpaParagraphs)
        {
            if (paragraph.ParagraphProperties?.Tabs?.Elements<TabStop>().Any() != true)
            {
                report.Issues.Add("落款矩阵注册会计师行缺少制表位");
                break;
            }
        }
    }

    private static void VerifySharedNumberingMatrix(WordprocessingDocument word, ScenarioVerificationReportContract report)
    {
        AssertHeading(word, report, "共享模板短标题", 3, true);

        var longBody = FindParagraph(word, "这是一个长度明显超过三十个字的长编号正文");
        if (longBody == null)
        {
            report.Issues.Add("共享模板矩阵缺少长编号正文");
            return;
        }

        var level = ResolveHeadingLevel(word, longBody);
        if (level != 0)
        {
            report.Issues.Add("共享模板矩阵中的长编号正文被误判成标题");
        }

        if (!HasNumbering(word, longBody))
        {
            report.Issues.Add("共享模板矩阵中的长编号正文丢失了编号");
        }
    }

    private static void VerifyHeaderFooterMatrix(WordprocessingDocument word, ScenarioVerificationReportContract report)
    {
        var sections = OpenXmlHelper.收集分节(word.MainDocumentPart?.Document?.Body);
        if (sections.Count < 2)
        {
            report.Issues.Add("页眉页脚分节矩阵缺少至少两个分节");
            return;
        }

        if (sections[0].Elements<HeaderReference>().Any() || sections[0].Elements<FooterReference>().Any())
        {
            report.Issues.Add("封面节页眉页脚仍然存在");
        }

        if (!sections[1].Elements<HeaderReference>().Any())
        {
            report.Issues.Add("正文节页眉被误删");
        }

        if (!sections[1].Elements<FooterReference>().Any())
        {
            report.Issues.Add("正文节页脚被误删");
        }
    }

    private static void VerifyRealSample(WordprocessingDocument word, ScenarioVerificationReportContract report)
    {
        AssertHeading(word, report, "公司的基本情况", 1, true);
        AssertHeading(word, report, "（一）基本情况", 2, true);
    }

    private static void VerifyTextBoxPitfallMatrix(WordprocessingDocument word, ScenarioVerificationReportContract report, FirmDocumentContext context)
    {
        var complex = 复杂结构观察服务.观察(word, context.HasCover);
        report.Facts["复杂结构.文本框数"] = complex.文本框数.ToString();
        if (complex.文本框数 < 1 || !complex.文本框摘要.Contains("文本框里的话", StringComparison.Ordinal))
        {
            report.Issues.Add("文本框坑点矩阵没有稳定保留文本框文字");
        }
    }

    private static void VerifyFootnoteEndnotePitfallMatrix(WordprocessingDocument word, ScenarioVerificationReportContract report, FirmDocumentContext context)
    {
        var complex = 复杂结构观察服务.观察(word, context.HasCover);
        report.Facts["复杂结构.脚注数"] = complex.脚注数.ToString();
        report.Facts["复杂结构.尾注数"] = complex.尾注数.ToString();
        if (complex.脚注数 < 1 || complex.尾注数 < 1)
        {
            report.Issues.Add("脚注尾注坑点矩阵丢失了脚注或尾注结构");
        }
    }

    private static void VerifyNumberingRestartPitfallMatrix(WordprocessingDocument word, ScenarioVerificationReportContract report, FirmDocumentContext context)
    {
        var complex = 复杂结构观察服务.观察(word, context.HasCover);
        report.Facts["复杂结构.编号重启数"] = complex.编号重启数.ToString();
        if (complex.编号重启数 < 1)
        {
            report.Issues.Add("编号重启坑点矩阵没有保留编号重启结构");
        }
    }

    private static void VerifyMergedCellPitfallMatrix(WordprocessingDocument word, ScenarioVerificationReportContract report, FirmDocumentContext context)
    {
        var complex = 复杂结构观察服务.观察(word, context.HasCover);
        report.Facts["复杂结构.横向合并数"] = complex.横向合并数.ToString();
        report.Facts["复杂结构.纵向合并数"] = complex.纵向合并数.ToString();
        if (complex.横向合并数 < 1 || complex.纵向合并数 < 1)
        {
            report.Issues.Add("合并单元格坑点矩阵丢失了横向或纵向合并结构");
        }
    }

    private static void VerifyComplexSignoffPitfallMatrix(WordprocessingDocument word, ScenarioVerificationReportContract report, FirmDocumentContext context)
    {
        var complex = 复杂结构观察服务.观察(word, context.HasCover);
        report.Facts["复杂结构.落款复杂结构数"] = complex.落款复杂结构数.ToString();
        if (complex.落款复杂结构数 < 2)
        {
            report.Issues.Add("落款复杂结构坑点矩阵没有保留落款区复杂结构");
        }
    }

    private static void AssertHeading(WordprocessingDocument word, ScenarioVerificationReportContract report, string text, int expectedLevel, bool expectNumbering)
    {
        var paragraph = FindParagraph(word, text);
        if (paragraph == null)
        {
            report.Issues.Add($"缺少段落：{text}");
            return;
        }

        var actualLevel = ResolveHeadingLevel(word, paragraph);
        if (actualLevel != expectedLevel)
        {
            report.Issues.Add($"段落“{text}”标题级别异常，期望 {expectedLevel}，实际 {actualLevel}");
        }

        if (expectNumbering && !HasNumbering(word, paragraph) && !HasVisibleHeadingPrefix(paragraph, expectedLevel))
        {
            report.Issues.Add($"段落“{text}”缺少编号语义");
        }
    }

    private static void AssertBodyNumberingPreserved(WordprocessingDocument word, ScenarioVerificationReportContract report, string text)
    {
        var paragraph = FindParagraph(word, text);
        if (paragraph == null)
        {
            report.Issues.Add($"缺少正文自动编号段落：{text}");
            return;
        }

        if (LooksLikeWrittenHeading(paragraph))
        {
            report.Issues.Add($"正文自动编号段落被误判成标题：{text}");
        }

        if (!HasNumbering(word, paragraph))
        {
            report.Issues.Add($"正文自动编号段落丢失编号：{text}");
        }
    }

    private static bool LooksLikeWrittenHeading(Paragraph paragraph)
    {
        var outline = paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value;
        if (outline is 0 or 1 or 2) return true;

        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.Trim().ToLowerInvariant();
        return styleId is "1" or "heading1" or "标题1" or "h1"
            or "2" or "heading2" or "标题2" or "h2"
            or "3" or "heading3" or "标题3" or "h3";
    }

    private static Paragraph? FindParagraph(WordprocessingDocument word, string containsText)
    {
        var body = word.MainDocumentPart?.Document?.Body;
        if (body == null) return null;

        return body.Descendants<Paragraph>()
            .FirstOrDefault(p => NormalizeText(GetVisibleText(p)).Contains(NormalizeText(containsText), StringComparison.Ordinal));
    }

    private static int ResolveHeadingLevel(WordprocessingDocument word, Paragraph paragraph)
    {
        var text = NormalizeText(GetVisibleText(paragraph));
        if (text.Length <= 30)
        {
            if (一级标题文本.IsMatch(text)) return 1;
            if (二级标题文本.IsMatch(text)) return 2;
            if (三级标题文本.IsMatch(text)) return 3;
        }

        var outline = paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value;
        if (outline == 0) return 1;
        if (outline == 1) return 2;
        if (outline == 2) return 3;

        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.Trim().ToLowerInvariant();
        if (styleId is "1" or "heading1" or "标题1" or "h1") return 1;
        if (styleId is "2" or "heading2" or "标题2" or "h2") return 2;
        if (styleId is "3" or "heading3" or "标题3" or "h3") return 3;

        return 0;
    }

    private static bool HasNumbering(WordprocessingDocument word, Paragraph paragraph)
    {
        var direct = paragraph.ParagraphProperties?.GetFirstChild<NumberingProperties>();
        if (direct != null) return true;

        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (string.IsNullOrWhiteSpace(styleId)) return false;

        var styles = word.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles == null) return false;

        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var current = styleId;
        while (!string.IsNullOrWhiteSpace(current) && visited.Add(current))
        {
            var style = styles.Elements<Style>().FirstOrDefault(x => x.StyleId?.Value == current);
            if (style?.StyleParagraphProperties?.GetFirstChild<NumberingProperties>() != null)
            {
                return true;
            }

            current = style?.BasedOn?.Val?.Value;
        }

        return false;
    }

    private static bool HasVisibleHeadingPrefix(Paragraph paragraph, int expectedLevel)
    {
        var text = NormalizeText(GetVisibleText(paragraph));
        return expectedLevel switch
        {
            1 => 一级标题文本.IsMatch(text),
            2 => 二级标题文本.IsMatch(text),
            3 => Regex.IsMatch(text, @"^\d{1,3}[、\.．]"),
            _ => false
        };
    }

    private static string GetVisibleText(OpenXmlElement? element)
    {
        return OpenXmlHelper.提取可见文本(element);
    }

    private static string NormalizeText(string? text)
    {
        return OpenXmlHelper.NormalizeText(text);
    }

    private static string NormalizeScenarioName(string scenarioName)
    {
        const string 已排版后缀 = "_已排版";
        var key = Path.GetFileNameWithoutExtension(scenarioName);
        return key.EndsWith(已排版后缀, StringComparison.OrdinalIgnoreCase)
            ? key[..^已排版后缀.Length]
            : key;
    }

    private static void AssertEquals(ScenarioVerificationReportContract report, string label, string expected, string actual)
    {
        if (!string.Equals(expected, actual, StringComparison.Ordinal))
        {
            report.Issues.Add($"{label}不一致，期望“{expected}”，实际“{actual}”");
        }
    }
}
