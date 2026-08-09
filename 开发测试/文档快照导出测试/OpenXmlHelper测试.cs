using DocumentFormat.OpenXml.Wordprocessing;
using FirmFormatter.OpenXml.Core;
using Xunit;

namespace 文档快照导出测试;

public sealed class OpenXmlHelper测试
{
    [Fact]
    public void 收集分节时_应依文档顺序返回段落分节和正文末尾分节()
    {
        var paragraphSection = new SectionProperties(new PageSize { Width = 100 });
        var bodySection = new SectionProperties(new PageSize { Width = 200 });
        var body = new Body(
            new Paragraph(new ParagraphProperties(paragraphSection), new Run(new Text("正文"))),
            bodySection);

        var sections = OpenXmlHelper.收集分节(body);

        Assert.Equal(2, sections.Count);
        Assert.Same(paragraphSection, sections[0]);
        Assert.Same(bodySection, sections[1]);
    }

    [Fact]
    public void 收集空正文分节时_应返回空集合()
    {
        Assert.Empty(OpenXmlHelper.收集分节(new Body()));
    }
}
