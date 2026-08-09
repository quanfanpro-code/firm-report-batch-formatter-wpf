using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FirmFormatter.OpenXml.Core;
using Xunit;

namespace 文档快照导出测试;

public sealed class 封面同节三级标题测试
{
    [Fact]
    public void 有封面且封面正文同节时_三级标题分隔符仍应规范为全角点()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocument(
            stream,
            new Paragraph(new Run(new Text("审计报告"))),
            new Paragraph(new Run(new Text("泊微公司"))),
            new Paragraph(new Run(new Text("1、审计报告"))));

        new ParagraphService().Apply(word, true, []);

        var heading = word.MainDocumentPart!
            .Document!
            .Body!
            .Elements<Paragraph>()
            .Last();

        Assert.Equal("1．审计报告", OpenXmlHelper.ParagraphText(heading).Trim());
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
