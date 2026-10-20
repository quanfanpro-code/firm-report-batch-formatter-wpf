using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FirmFormatter.OpenXml.Contracts;
using FirmFormatter.OpenXml.Core;
using Xunit;

namespace 文档快照导出测试;

// 复现缺陷：源稿把大标题挂在"标题 1"样式上，排版侧按大标题规格排（居中、加粗、大字），
// 门禁却沿样式把它解析成一级标题，再按"小四加粗 + 写大纲级别"的章节标题规格拦截，形成排版-门禁死锁。
// 修复口径（2026-10-20 用户拍板）：大标题不是章节标题，排版时摘掉使其被解析为章节标题的样式挂靠；
// 字号统一小二（36 半点），行距、段前段后、居中、缩进保持大标题既有排法。
public sealed class 大标题样式摘除测试
{
    [Fact]
    public void 大标题挂标题1样式_排版后必须摘除章节标题样式挂靠()
    {
        using var stream = new MemoryStream();
        using var word = 创建访谈记录式文档(stream);

        new ParagraphService().Apply(word, false, []);

        var title = word.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
        Assert.Null(title.ParagraphProperties?.ParagraphStyleId);
    }

    [Fact]
    public void 大标题挂标题1样式_排版后门禁不得按章节标题规格误拦()
    {
        using var stream = new MemoryStream();
        using var word = 创建访谈记录式文档(stream);
        var request = new RequestContract { ScenarioName = "常规" };
        var context = new FirmDocumentClassifier().BuildContext(word, request);
        new HeaderFooterService().Apply(word, context.HasCover);
        var signoff = new SignoffService().IdentifySignoffParagraphs(word, context.HasCover);
        new ParagraphService().Apply(word, context.HasCover, signoff);
        new TableService().Apply(word, context.HasCover);
        new SignoffService().Apply(word, signoff);

        var gate = new GateCheckService().Run(word, request, context);

        var 标题类问题 = gate.BlockingIssues.Where(i => i.Code is "heading_layout" or "heading_font").ToList();
        Assert.True(标题类问题.Count == 0,
            "大标题被按章节标题规格误拦：" + string.Join("；", 标题类问题.Select(i => $"{i.Code}：{i.Message}")));
    }

    // 复刻真实样本形态：大标题挂 Heading1（样式自带大纲级别），正文长段落加末节属性
    private static WordprocessingDocument 创建访谈记录式文档(MemoryStream stream)
    {
        var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var mainPart = word.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(
            new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                new Run(new RunProperties(new Bold(), new FontSize { Val = "36" }), new Text("访谈记录"))),
            new Paragraph(new Run(new Text(new string('正', 260)))),
            new SectionProperties()));
        mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
            new Style(
                new StyleName { Val = "heading 1" },
                new StyleParagraphProperties(new OutlineLevel { Val = 0 }))
            {
                Type = StyleValues.Paragraph,
                StyleId = "Heading1"
            });
        mainPart.Document.Save();
        return word;
    }
}
