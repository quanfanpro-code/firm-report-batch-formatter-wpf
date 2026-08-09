using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using FirmFormatter.OpenXml.Core;
using Xunit;

namespace 文档快照导出测试;

public sealed class 可见文本抽取测试
{
    [Fact]
    public void 提取可见文本时_应忽略隐藏文字()
    {
        var paragraph = new Paragraph(
            new Run(new Text("可见")),
            new Run(
                new RunProperties(new Vanish()),
                new Text("隐藏")),
            new Run(new Text("文本")));

        var actual = OpenXmlHelper.提取可见文本(paragraph);
        Assert.Equal("可见文本", actual);
    }

    [Fact]
    public void 提取可见文本时_应保留内容控件里的文字()
    {
        var paragraph = new Paragraph(
            new SdtRun(
                new SdtProperties(),
                new SdtContentRun(
                    new Run(new Text("控件内容")))));

        var actual = OpenXmlHelper.提取可见文本(paragraph);
        Assert.Equal("控件内容", actual);
    }

    [Fact]
    public void 提取可见文本时_应跳过Drawing里的脏文本()
    {
        var drawing = new Drawing();
        drawing.Append(new Text("锚点脏文本"));

        var paragraph = new Paragraph(
            new Run(new Text("前缀")),
            new Run(drawing),
            new Run(new Text("后缀")));

        var actual = OpenXmlHelper.提取可见文本(paragraph);
        Assert.Equal("前缀后缀", actual);
    }

    [Fact]
    public void 提取文本框文本时_应单独读出TextBoxContent里的文字()
    {
        var drawing = new Drawing();
        drawing.Append(new TextBoxContent(
            new Paragraph(
                new Run(new Text("文本框内容")))));

        var paragraph = new Paragraph(
            new Run(new Text("正文")),
            new Run(drawing));

        Assert.Equal("正文", OpenXmlHelper.提取可见文本(paragraph));
        Assert.Equal("文本框内容", OpenXmlHelper.提取文本框文本(paragraph));
    }
}
