using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FirmFormatter.OpenXml.Core;
using Xunit;

namespace 文档快照导出测试;

// 复现缺陷：大标题识别写死"最多 3 行"，4 行标题的报告第 4 行被降级成正文。
// 修复后：只要下一行仍独立满足大标题判据（加粗且大于小四）就继续纳入，安全上限 6 行。
public sealed class 大标题行数自适应测试
{
    [Fact]
    public void 四行加大加粗标题_第四行也应按大标题排版()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocument(stream,
            标题段("宜宾港建商业运营管理有限责任公司"),
            标题段("宜宾三江新区高铁片区保障性租赁住房建设项目截至"),
            标题段("2026年8月7日资本金到位及使用情况"),
            标题段("专项审核报告"),
            正文段("这里是正文内容，应当保持正文格式。"));

        new ParagraphService().Apply(word, false, []);

        var paragraphs = word.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();
        var fourth = paragraphs[3];
        Assert.Equal(JustificationValues.Center, fourth.ParagraphProperties!.Justification!.Val!.Value);
        var fourthRun = fourth.Elements<Run>().Single(r => !string.IsNullOrWhiteSpace(r.InnerText));
        Assert.Equal("32", fourthRun.RunProperties!.FontSize!.Val!.Value);

        // 第五行是普通正文：标题区必须在第四行后正常结束，不能误吞
        var fifthRun = paragraphs[4].Elements<Run>().First(r => !string.IsNullOrWhiteSpace(r.InnerText));
        Assert.Equal("24", fifthRun.RunProperties!.FontSize!.Val!.Value);
    }

    [Fact]
    public void 超过六行的加大加粗段落_第七行起不再算标题()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocument(stream,
            标题段("标题一"), 标题段("标题二"), 标题段("标题三"),
            标题段("标题四"), 标题段("标题五"), 标题段("标题六"),
            标题段("标题七"));

        new ParagraphService().Apply(word, false, []);

        var seventh = word.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList()[6];
        Assert.NotEqual(JustificationValues.Center, seventh.ParagraphProperties!.Justification?.Val?.Value);
    }

    private static Paragraph 标题段(string text) =>
        new(new Run(new RunProperties(new Bold(), new FontSize { Val = "32" }), new Text(text)));

    private static Paragraph 正文段(string text) =>
        new(new Run(new RunProperties(new FontSize { Val = "24" }), new Text(text)));

    private static WordprocessingDocument CreateDocument(MemoryStream stream, params OpenXmlElement[] bodyChildren)
    {
        var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var mainPart = word.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(bodyChildren));
        mainPart.Document.Save();
        return word;
    }
}
