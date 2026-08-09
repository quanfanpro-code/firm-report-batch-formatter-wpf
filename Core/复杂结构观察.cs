using System.Text;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FirmFormatter.OpenXml.Core;

public sealed class 复杂结构观察结果
{
    public required int 文本框数 { get; init; }
    public required int 脚注数 { get; init; }
    public required int 尾注数 { get; init; }
    public required int 编号重启数 { get; init; }
    public required int 横向合并数 { get; init; }
    public required int 纵向合并数 { get; init; }
    public required int 书签数 { get; init; }
    public required int 超链接数 { get; init; }
    public required int 域代码数 { get; init; }
    public required int 落款复杂结构数 { get; init; }
    public required string 文本框摘要 { get; init; }
    public required string 脚注尾注摘要 { get; init; }
    public required string 文本框内容指纹 { get; init; }
    public required string 脚注尾注内容指纹 { get; init; }
}

public static class 复杂结构观察服务
{
    private const int 摘要最大长度 = 500;

    public static 复杂结构观察结果 观察(
        WordprocessingDocument word,
        bool hasCover = false)
    {
        var body = word.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("文档缺少正文");

        var textBoxes = body.Descendants<Paragraph>()
            .Select(paragraph => OpenXmlHelper.提取文本框文本(paragraph))
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToList();
        var footnotes = 收集脚注文字(word.MainDocumentPart?.FootnotesPart);
        var endnotes = 收集尾注文字(word.MainDocumentPart?.EndnotesPart);
        var signoffParagraphs = new SignoffService().IdentifySignoffParagraphs(word, hasCover);

        return new 复杂结构观察结果
        {
            文本框数 = textBoxes.Count,
            脚注数 = footnotes.Count,
            尾注数 = endnotes.Count,
            编号重启数 = 收集编号重启数(word),
            横向合并数 = body.Descendants<GridSpan>().Count(span => (span.Val?.Value ?? 1) > 1),
            纵向合并数 = body.Descendants<VerticalMerge>().Count(),
            书签数 = body.Descendants<BookmarkStart>().Count(),
            超链接数 = body.Descendants<Hyperlink>().Count(),
            域代码数 = body.Descendants<FieldCode>().Count(),
            落款复杂结构数 = 收集落款复杂结构数(signoffParagraphs),
            文本框摘要 = 拼接摘要(textBoxes),
            脚注尾注摘要 = 拼接摘要(footnotes.Concat(endnotes)),
            文本框内容指纹 = 计算内容指纹(textBoxes),
            脚注尾注内容指纹 = 计算内容指纹(footnotes.Concat(endnotes))
        };
    }

    private static List<string> 收集脚注文字(FootnotesPart? footnotesPart)
    {
        if (footnotesPart?.Footnotes == null)
        {
            return [];
        }

        return footnotesPart.Footnotes.Elements<Footnote>()
            .Where(note => note.Type == null && (note.Id?.Value ?? 0) >= 0)
            .Select(OpenXmlHelper.提取可见文本)
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToList();
    }

    private static List<string> 收集尾注文字(EndnotesPart? endnotesPart)
    {
        if (endnotesPart?.Endnotes == null)
        {
            return [];
        }

        return endnotesPart.Endnotes.Elements<Endnote>()
            .Where(note => note.Type == null && (note.Id?.Value ?? 0) >= 0)
            .Select(OpenXmlHelper.提取可见文本)
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToList();
    }

    private static string 拼接摘要(IEnumerable<string> texts)
    {
        var result = new StringBuilder();
        foreach (var text in texts.Select(OpenXmlHelper.NormalizeText).Where(text => !string.IsNullOrWhiteSpace(text)))
        {
            if (result.Length > 0)
            {
                result.Append(" | ");
            }

            result.Append(text);
        }

        // 摘要仅用于界面展示，截断防止超长文本撑爆界面
        var summary = result.ToString();
        return summary.Length > 摘要最大长度
            ? summary[..摘要最大长度] + "…"
            : summary;
    }

    private static string 计算内容指纹(IEnumerable<string> texts)
    {
        var normalized = string.Join("\n", texts.Select(OpenXmlHelper.NormalizeText));
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(normalized)));
    }

    private static int 收集编号重启数(WordprocessingDocument word)
    {
        return word.MainDocumentPart?.NumberingDefinitionsPart?.Numbering?
            .Descendants<LevelOverride>()
            .Count(item => item.GetFirstChild<StartOverrideNumberingValue>() != null) ?? 0;
    }

    private static int 收集落款复杂结构数(IEnumerable<Paragraph> signoffParagraphs)
    {
        return signoffParagraphs.Count(段落包含复杂结构);
    }

    private static bool 段落包含复杂结构(Paragraph paragraph)
    {
        return paragraph.Descendants<BookmarkStart>().Any()
            || paragraph.Descendants<Hyperlink>().Any()
            || paragraph.Descendants<FieldCode>().Any()
            || !string.IsNullOrWhiteSpace(OpenXmlHelper.提取文本框文本(paragraph));
    }
}
