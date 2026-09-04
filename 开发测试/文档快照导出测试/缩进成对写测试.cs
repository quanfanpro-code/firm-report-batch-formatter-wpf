using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FirmFormatter.OpenXml.Core;
using Xunit;

namespace 文档快照导出测试;

// 复现缺陷：程序只写缩进缇值、不写字符单位时，样式链上的 FirstLineChars 会架空缇值
//（落款台阶塌平、章节标题被多缩进 2 字符）。修复后两者必须成对出现。
public sealed class 缩进成对写测试
{
    [Fact]
    public void 落款台阶行_缇值与字符单位必须成对()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocument(stream,
            new Paragraph(new Run(new Text("四川华信(集团)会计师事务所"))),
            new Paragraph(new Run(new Text("（特殊普通合伙）"))),
            new Paragraph(new Run(new Text("中国·成都"))),
            new Paragraph(new Run(new Text("二〇二六年八月十五日"))));

        var paragraphs = word.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();
        new SignoffService().Apply(word, paragraphs);

        var ind0 = paragraphs[0].ParagraphProperties!.Indentation!;
        Assert.Equal("0", ind0.FirstLine?.Value);
        Assert.Equal(0, ind0.FirstLineChars?.Value);
        var ind1 = paragraphs[1].ParagraphProperties!.Indentation!;
        Assert.Equal("480", ind1.FirstLine?.Value);
        Assert.Equal(200, ind1.FirstLineChars?.Value);
        var ind2 = paragraphs[2].ParagraphProperties!.Indentation!;
        Assert.Equal("960", ind2.FirstLine?.Value);
        Assert.Equal(400, ind2.FirstLineChars?.Value);
    }

    [Fact]
    public void 落款日期行_首行缩进字符单位必须清零()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocument(stream,
            new Paragraph(new Run(new Text("四川华信(集团)会计师事务所"))),
            new Paragraph(new Run(new Text("二〇二六年八月十五日"))));

        var paragraphs = word.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();
        new SignoffService().Apply(word, paragraphs);

        var ind = paragraphs[1].ParagraphProperties!.Indentation!;
        Assert.Equal("0", ind.FirstLine?.Value);
        Assert.Equal(0, ind.FirstLineChars?.Value);
    }

    [Fact]
    public void 一级标题_首行缩进字符单位必须清零()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocument(stream,
            new Paragraph(new Run(new Text("一、项目资本金范围"))));

        new ParagraphService().Apply(word, false, []);

        var ind = word.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Single()
            .ParagraphProperties!.Indentation!;
        Assert.Equal("0", ind.FirstLine?.Value);
        Assert.Equal(0, ind.FirstLineChars?.Value);
    }

    [Fact]
    public void 三级标题_保持两字符首行缩进()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocument(stream,
            new Paragraph(new Run(new Text("1．基本情况"))));

        new ParagraphService().Apply(word, false, []);

        var ind = word.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Single()
            .ParagraphProperties!.Indentation!;
        Assert.Equal("480", ind.FirstLine?.Value);
        Assert.Equal(200, ind.FirstLineChars?.Value);
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
