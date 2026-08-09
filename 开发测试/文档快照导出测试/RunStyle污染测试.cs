using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FirmFormatter.OpenXml.Core;
using Xunit;

namespace 文档快照导出测试;

public sealed class RunStyle污染测试
{
    [Fact]
    public void 正文格式化后_应清掉RunStyle污染()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocument(
            stream,
            new Paragraph(
                new Run(
                    new RunProperties(new RunStyle { Val = "脏字符样式" }),
                    new Text("这是正文内容"))));

        new ParagraphService().Apply(word, false, []);

        var paragraph = word.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single();
        var run = paragraph.Elements<Run>().Single();

        Assert.Null(run.RunProperties?.RunStyle);
    }

    [Fact]
    public void 落款格式化后_应清掉普通行RunStyle污染()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocument(
            stream,
            new Paragraph(
                new Run(
                    new RunProperties(new RunStyle { Val = "脏字符样式" }),
                    new Text("四川华信(集团)会计师事务所"))),
            new Paragraph(new Run(new Text("2025年5月2日"))));

        var paragraphs = word.MainDocumentPart!.Document.Body!.Elements<Paragraph>().ToList();
        new SignoffService().Apply(word, paragraphs);

        var run = paragraphs[0].Elements<Run>().Single();
        Assert.Null(run.RunProperties?.RunStyle);
    }

    private static WordprocessingDocument CreateDocument(MemoryStream stream, params OpenXmlElement[] bodyChildren)
    {
        var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var mainPart = word.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(bodyChildren));
        mainPart.Document.Save();
        return word;
    }
}
