using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using FirmFormatter.OpenXml.Contracts;

namespace FirmFormatter.OpenXml.Core;

public sealed class ValidationService
{
    private static readonly Regex 三级标题文本 = new(@"^\d{1,3}．", RegexOptions.Compiled);
    private readonly FirmRuleProfile _ruleProfile = FirmRuleProfile.Default;

    public ValidationReportContract Validate(
        WordprocessingDocument word,
        bool hasCover = false,
        bool isPureCoverDocument = false,
        string? scenarioName = null,
        bool throwOnFailure = true)
    {
        var report = new ValidationReportContract();
        var main = word.MainDocumentPart;
        var body = main?.Document?.Body;

        if (main == null || body == null)
        {
            report.AddIssue("结构", "missing_body", "缺少正文");
            if (throwOnFailure)
            {
                throw new InvalidDataException(BuildFailureMessage(report));
            }

            return report;
        }

        report.SetFact("可见正文字符数", OpenXmlHelper.统计正文可见有效文本字数(body).ToString());
        report.SetFact("场景", string.IsNullOrWhiteSpace(scenarioName) ? "常规" : scenarioName);

        ValidateStructure(word, body, report);
        ValidateBusiness(word, body, hasCover, isPureCoverDocument, report);

        if (!report.Success && throwOnFailure)
        {
            throw new InvalidDataException(BuildFailureMessage(report));
        }

        return report;
    }

    private static void ValidateStructure(WordprocessingDocument word, Body body, ValidationReportContract report)
    {
        if (!OpenXmlHelper.正文包含有效可见内容(body))
        {
            report.AddIssue("结构", "empty_body", "正文为空");
        }

        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        foreach (var error in validator.Validate(word).Take(50))
        {
            // OpenXmlValidator 的 schema 排序差异不阻断流水线，只记警告
            report.AddWarning("结构", "openxml_validator", $"{error.Path?.XPath ?? "未知路径"}：{error.Description}");
        }

        var bodySectionProperties = body.Elements<SectionProperties>().ToList();
        if (bodySectionProperties.Count > 1)
        {
            report.AddIssue("结构", "multiple_body_sectpr", "Body 下存在多个直属 SectionProperties");
        }

        if (bodySectionProperties.Count == 1 && !ReferenceEquals(body.LastChild, bodySectionProperties[0]))
        {
            report.AddIssue("结构", "body_sectpr_not_last", "Body 直属 SectionProperties 不是最后一个子节点");
        }

        foreach (var cell in body.Descendants<TableCell>())
        {
            if (!cell.Elements<Paragraph>().Any())
            {
                report.AddIssue("结构", "cell_without_paragraph", "存在不含段落的表格单元格");
            }
        }

        var main = word.MainDocumentPart;
        if (main == null) return;

        foreach (var section in OpenXmlHelper.收集分节(body))
        {
            foreach (var headerReference in section.Elements<HeaderReference>())
            {
                if (string.IsNullOrWhiteSpace(headerReference.Id?.Value))
                {
                    report.AddIssue("结构", "header_ref_missing_id", "页眉引用缺少关系 Id");
                    continue;
                }

                // 用 Parts 探测代替 GetPartById 裸 catch，避免把无关异常也归因为引用损坏
                if (!main.Parts.Any(part => part.RelationshipId == headerReference.Id.Value))
                {
                    report.AddIssue("结构", "header_ref_broken", $"页眉引用损坏：{headerReference.Id.Value}");
                }
            }

            foreach (var footerReference in section.Elements<FooterReference>())
            {
                if (string.IsNullOrWhiteSpace(footerReference.Id?.Value))
                {
                    report.AddIssue("结构", "footer_ref_missing_id", "页脚引用缺少关系 Id");
                    continue;
                }

                if (!main.Parts.Any(part => part.RelationshipId == footerReference.Id.Value))
                {
                    report.AddIssue("结构", "footer_ref_broken", $"页脚引用损坏：{footerReference.Id.Value}");
                }
            }
        }
    }

    private void ValidateBusiness(
        WordprocessingDocument word,
        Body body,
        bool hasCover,
        bool isPureCoverDocument,
        ValidationReportContract report)
    {
        var main = word.MainDocumentPart;
        if (main == null) return;

        ValidateHeadingSemantics(main, body, report);
        ValidateAutomaticBodyNumbering(main, body, report);
        ValidateTableRules(body, report);
        ValidateSectionMargins(body, isPureCoverDocument, report, _ruleProfile);
        ValidateSignoff(word, hasCover, report);
    }

    private static void ValidateHeadingSemantics(MainDocumentPart mainPart, Body body, ValidationReportContract report)
    {
        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            if (paragraph.Ancestors<Table>().Any()) continue;

            var text = OpenXmlHelper.NormalizeText(OpenXmlHelper.ParagraphText(paragraph));
            if (string.IsNullOrWhiteSpace(text)) continue;

            var level = OpenXmlHelper.ResolveHeadingLevelForValidation(mainPart, paragraph);
            if (level == 0) continue;

            if (level == 3
                && !三级标题文本.IsMatch(text)
                && text.Length <= 30
                && Regex.IsMatch(text, @"^\d{1,3}[、\.]"))
            {
                report.AddIssue("业务", "h3_separator_not_normalized", $"三级标题分隔符未统一为全角点：{text}");
            }
        }
    }

    private static void ValidateAutomaticBodyNumbering(MainDocumentPart mainPart, Body body, ValidationReportContract report)
    {
        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            if (paragraph.Ancestors<Table>().Any()) continue;

            var text = OpenXmlHelper.NormalizeText(OpenXmlHelper.ParagraphText(paragraph));
            if (string.IsNullOrWhiteSpace(text)) continue;

            if (OpenXmlHelper.ResolveHeadingLevelForValidation(mainPart, paragraph) != 0) continue;
            if (!OpenXmlHelper.段落存在编号(mainPart, paragraph)) continue;

            if (OpenXmlHelper.是正文式四级(mainPart, paragraph))
            {
                continue;
            }

            var source = OpenXmlHelper.Resolve编号来源(mainPart, paragraph);
            if (source == "无")
            {
                report.AddIssue("业务", "body_numbering_missing", $"正文自动编号已丢失：{text}");
            }
        }
    }

    private static void ValidateTableRules(Body body, ValidationReportContract report)
    {
        foreach (var table in body.Elements<Table>())
        {
            var rows = table.Elements<TableRow>().ToList();
            for (var rowIndex = 1; rowIndex < rows.Count; rowIndex++)
            {
                var cells = rows[rowIndex].Elements<TableCell>().ToList();
                if (cells.Count == 0) continue;

                var firstCellText = OpenXmlHelper.NormalizeText(OpenXmlHelper.提取可见文本(cells[0]));
                // 先用更宽的正则圈定数字样文本，再分别判断对齐与格式，与 TableService 首列整数格式化行为对齐
                if (Regex.IsMatch(firstCellText, @"^[\d.,]+$"))
                {
                    var justification = cells[0].Elements<Paragraph>()
                        .FirstOrDefault()?.ParagraphProperties?.Justification?.Val?.Value;
                    if (justification != JustificationValues.Left)
                    {
                        report.AddIssue("业务", "table_first_col_alignment", $"表格首列正整数字符串未左对齐：{firstCellText}");
                    }

                    if (firstCellText.Contains('.') || firstCellText.Contains(','))
                    {
                        report.AddIssue("业务", "table_first_col_integer_format", $"表格首列正整数字符串被错误格式化：{firstCellText}");
                    }
                }
            }
        }
    }

    private static void ValidateSectionMargins(Body body, bool isPureCoverDocument, ValidationReportContract report, FirmRuleProfile ruleProfile)
    {
        foreach (var section in OpenXmlHelper.收集分节(body))
        {
            var pageMargin = section.GetFirstChild<PageMargin>();
            if (pageMargin == null)
            {
                report.AddIssue("业务", "missing_page_margin", "分节缺少页边距定义");
                continue;
            }

            if (pageMargin.Top?.Value != (uint)ruleProfile.页边距上Twips || pageMargin.Left?.Value != (uint)ruleProfile.页边距左Twips)
            {
                report.AddIssue("业务", "page_margin_top_left", "页边距未收口到目标值");
            }

            if (isPureCoverDocument && pageMargin.Bottom?.Value != (uint)ruleProfile.纯封面下边距Twips)
            {
                report.AddIssue("业务", "pure_cover_bottom_margin", $"纯封面下边距未收口到 {ruleProfile.纯封面下边距Twips / 567.0:0.0}cm");
            }
        }
    }

    private void ValidateSignoff(
        WordprocessingDocument word,
        bool hasCover,
        ValidationReportContract report)
    {
        var signoffService = new SignoffService();
        var signoffParagraphs = signoffService.IdentifySignoffParagraphs(word, hasCover);
        if (signoffParagraphs.Count == 0) return;

        var cpaCount = signoffParagraphs.Count(paragraph =>
            OpenXmlHelper.NormalizeText(OpenXmlHelper.ParagraphText(paragraph)).Contains("中国注册会计师", StringComparison.Ordinal));
        if (cpaCount < 2)
        {
            report.AddIssue("业务", "signoff_cpa_count", "落款区内注册会计师行少于 2 行");
        }
    }

    private static string BuildFailureMessage(ValidationReportContract report)
    {
        if (report.Success) return "验证通过";

        var builder = new StringBuilder("验证失败：");
        foreach (var issue in report.Issues.Take(8))
        {
            builder.AppendLine();
            builder.Append($"- [{issue.Layer}] {issue.Code}：{issue.Message}");
        }

        if (report.Issues.Count > 8)
        {
            builder.AppendLine();
            builder.Append($"- 其余 {report.Issues.Count - 8} 项问题已省略");
        }

        return builder.ToString();
    }
}
