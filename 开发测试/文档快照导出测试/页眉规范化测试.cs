using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FirmFormatter.OpenXml.Core;
using Xunit;

namespace 文档快照导出测试;

// 复现缺陷：页眉用"一长串空格"把右侧文字顶到行尾，空格宽度随字号变化，
// 程序把页眉统一放大到五号后整行超出版心、右侧文字被挤到第二行。
// 修复后：空格填充转换为右对齐定位点，行首空白/制表符清除，字号变化不再影响布局。
public sealed class 页眉规范化测试
{
    [Fact]
    public void 空格凑右对齐的页眉_应转成右定位点()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocumentWithHeader(stream,
            new Paragraph(
                new Run(new TabChar()),
                new Run(new Text("宜宾港建商业运营管理有限责任公司")),
                new Run(new Text("                    ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new Text("专项审核报告"))));

        new HeaderFooterService().Apply(word, hasCover: false);

        var p = GetOnlyHeader(word).Descendants<Paragraph>().Single();
        var text = string.Concat(p.Descendants<Text>().Select(t => t.Text));
        Assert.Equal("宜宾港建商业运营管理有限责任公司专项审核报告", text);
        Assert.Single(p.Descendants<TabChar>());
        var tab = p.ParagraphProperties!.Tabs!.Elements<TabStop>().Single();
        Assert.Equal(TabStopValues.Right, tab.Val!.Value);
        Assert.Equal(9639, tab.Position!.Value); // 版心宽 = 11906-1417-850
    }

    [Fact]
    public void 单个空格的普通页眉_不转换()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocumentWithHeader(stream,
            new Paragraph(new Run(new Text("公司 报告"))));

        new HeaderFooterService().Apply(word, hasCover: false);

        var p = GetOnlyHeader(word).Descendants<Paragraph>().Single();
        Assert.Null(p.ParagraphProperties!.Tabs);
        Assert.Empty(p.Descendants<TabChar>());
    }

    [Fact]
    public void 含图片的页眉段落_跳过转换()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocumentWithHeader(stream,
            new Paragraph(
                new Run(new Text("公司        报告") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new Drawing())));

        new HeaderFooterService().Apply(word, hasCover: false);

        var p = GetOnlyHeader(word).Descendants<Paragraph>().Single();
        Assert.NotEmpty(p.Descendants<Drawing>());
        Assert.Null(p.ParagraphProperties!.Tabs);
    }

    private static WordprocessingDocument CreateDocumentWithHeader(MemoryStream stream, Paragraph headerParagraph)
    {
        var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var main = word.AddMainDocumentPart();
        var headerPart = main.AddNewPart<HeaderPart>();
        headerPart.Header = new Header(headerParagraph);
        headerPart.Header.Save();
        var sectPr = new SectionProperties(new HeaderReference
        {
            Type = HeaderFooterValues.Default,
            Id = main.GetIdOfPart(headerPart)
        });
        main.Document = new Document(new Body(new Paragraph(new Run(new Text("正文"))), sectPr));
        main.Document.Save();
        return word;
    }

    private static Header GetOnlyHeader(WordprocessingDocument word) =>
        word.MainDocumentPart!.HeaderParts.Single().Header!;
}
