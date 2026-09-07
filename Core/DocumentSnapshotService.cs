using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FirmFormatter.OpenXml.Contracts;

namespace FirmFormatter.OpenXml.Core;

public sealed class DocumentSnapshotService
{
    public string ExportAsJson(string docxPath)
    {
        using var word = WordprocessingDocument.Open(docxPath, false);
        var classifier = new FirmDocumentClassifier();
        var context = classifier.BuildContext(word, new RequestContract
        {
            InputPath = docxPath,
            // 快照导出是只读场景，不产生输出文件，输出路径置空以免被下游误当作真实输出
            OutputPath = string.Empty
        });
        return ExportAsJson(word, docxPath, context);
    }

    public string ExportAsJson(WordprocessingDocument word, string docxPath, FirmDocumentContext context)
    {
        var body = word.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("文档缺少正文");
        var sections = OpenXmlHelper.收集分节(body);
        var headings = CollectHeadings(word);
        var tables = CollectTables(body);
        var complex = context.复杂结构;
        var signoffParagraphs = new SignoffService()
            .IdentifySignoffParagraphs(word, context.HasCover);

        var payload = new
        {
            文档路径 = docxPath,
            正文可见字数 = context.VisibleTextLength,
            段落数 = body.Descendants<Paragraph>().Count(),
            表格数 = body.Descendants<Table>().Count(),
            表格详情 = tables,
            分节数 = sections.Count,
            标题数 = headings.Count,
            标题详情 = headings,
            分节详情 = sections.Select((section, index) => BuildSectionSnapshot(section, index + 1)).ToList(),
            落款命中数 = signoffParagraphs.Count,
            复杂结构 = complex
        };

        return JsonSerializer.Serialize(payload, new JsonSerializerOptions
        {
            WriteIndented = true
        });
    }

    private static List<object> CollectHeadings(WordprocessingDocument word)
    {
        var body = word.MainDocumentPart?.Document?.Body;
        if (body == null) return [];

        return body.Descendants<Paragraph>()
            .Select(paragraph => new
            {
                级别 = OpenXmlHelper.ResolveHeadingLevelForValidation(word.MainDocumentPart, paragraph),
                文本 = NormalizeVisibleText(GetVisibleText(paragraph))
            })
            .Where(item => item.级别 > 0 && !string.IsNullOrWhiteSpace(item.文本))
            .Cast<object>()
            .ToList();
    }

    private static List<object> CollectTables(Body body)
    {
        return body.Descendants<Table>()
            .Select((table, index) => new
            {
                序号 = index + 1,
                行数 = table.Elements<TableRow>().Count(),
                最大列数 = table.Elements<TableRow>().Select(row => row.Elements<TableCell>().Count()).DefaultIfEmpty(0).Max(),
                横向合并数 = table.Descendants<GridSpan>().Count(span => (span.Val?.Value ?? 1) > 1),
                纵向合并数 = table.Descendants<VerticalMerge>().Count()
            })
            .Cast<object>()
            .ToList();
    }

    private static string GetVisibleText(Paragraph paragraph)
    {
        return OpenXmlHelper.提取可见文本(paragraph);
    }

    private static string NormalizeVisibleText(string? text)
    {
        return OpenXmlHelper.NormalizeText(text);
    }

    private static object BuildSectionSnapshot(SectionProperties section, int index)
    {
        var pageSize = section.GetFirstChild<PageSize>();
        var pageMargin = section.GetFirstChild<PageMargin>();
        var isLandscape = pageSize?.Orient?.Value == PageOrientationValues.Landscape
            || (pageSize?.Width?.Value ?? 0) > (pageSize?.Height?.Value ?? 0);

        return new
        {
            序号 = index,
            方向 = isLandscape ? "横向" : "纵向",
            宽度 = (long?)pageSize?.Width?.Value ?? 0,
            高度 = (long?)pageSize?.Height?.Value ?? 0,
            页眉引用数 = section.Elements<HeaderReference>().Count(),
            页脚引用数 = section.Elements<FooterReference>().Count(),
            页边距 = new
            {
                上 = (long?)pageMargin?.Top?.Value ?? 0,
                下 = (long?)pageMargin?.Bottom?.Value ?? 0,
                左 = (long?)pageMargin?.Left?.Value ?? 0,
                右 = (long?)pageMargin?.Right?.Value ?? 0,
                页眉 = (long?)pageMargin?.Header?.Value ?? 0,
                页脚 = (long?)pageMargin?.Footer?.Value ?? 0
            }
        };
    }
}
