using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FirmFormatter.OpenXml.Contracts;
using FirmFormatter.OpenXml.Core;
using Xunit;
using V = DocumentFormat.OpenXml.Vml;

namespace 文档快照导出测试;

// 门禁与排版侧口径一致性回归测试：
// 凡是排版侧明确不改的区域（目录段落、文本框、落款区、纯封面文档、死页眉引用），
// 门禁不得再按标题语义或引用完整性拦截，否则形成"排版永不修、门禁必拦"的死锁。
// 背景：真实样本曾因"4、经营情况"被门禁拦截，而排版侧因文本带空白放弃替换。
public sealed class 门禁排版口径一致测试
{
    private static RequestContract BuildRequest() => new()
    {
        InputPath = "输入.docx",
        OutputPath = "输出.docx",
        RunSource = "测试"
    };

    // 与 DocumentPipeline 相同的排版分支：先分类，再按纯封面/常规路径排版，最后统一门禁
    private static GateCheckResultContract 排版后过门禁(WordprocessingDocument word)
    {
        var request = BuildRequest();
        var context = new FirmDocumentClassifier().BuildContext(word, request);

        if (context.IsPureCoverDocument)
        {
            new HeaderFooterService().ApplyPureCoverLayout(word);
        }
        else
        {
            new HeaderFooterService().Apply(word, context.HasCover);
            var signoff = new SignoffService().IdentifySignoffParagraphs(word, context.HasCover);
            new ParagraphService().Apply(word, context.HasCover, signoff);
            new TableService().Apply(word, context.HasCover);
        }

        return new GateCheckService().Run(word, request, context);
    }

    private static WordprocessingDocument CreateDocument(MemoryStream stream, params OpenXmlElement[] bodyChildren)
    {
        var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var mainPart = word.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(bodyChildren));
        mainPart.Document.Save();
        return word;
    }

    private static Paragraph 正文填充段() => new(new Run(new Text(new string('正', 260))));

    private static string 拼接问题(GateCheckResultContract gate)
        => string.Join("；", gate.BlockingIssues.Select(issue => $"[{issue.Layer}] {issue.Code}：{issue.Message}"));

    [Fact]
    public void 目录段落里的数字点号条目_门禁不应拦截()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocument(stream,
            new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "TOC1" }),
                new Run(new Text("1.经营情况"))),
            正文填充段(),
            new SectionProperties());

        var gate = 排版后过门禁(word);

        Assert.DoesNotContain(gate.BlockingIssues, issue => issue.Code == "h3_separator_not_normalized");
        Assert.True(gate.Success, 拼接问题(gate));
    }

    [Fact]
    public void 文本框段落里的数字点号文本_门禁不应拦截()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocument(stream,
            CreateTextBoxParagraph("1.备注说明"),
            正文填充段(),
            new SectionProperties());

        var gate = 排版后过门禁(word);

        Assert.DoesNotContain(gate.BlockingIssues, issue => issue.Code == "h3_separator_not_normalized");
        Assert.True(gate.Success, 拼接问题(gate));
    }

    [Fact]
    public void 落款区段落里的顿号编号行_门禁不应拦截()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocument(stream,
            正文填充段(),
            new Paragraph(new Run(new Text("四川华信(集团)会计师事务所"))),
            new Paragraph(new Run(new Text("1、联系地址：成都市高新区"))),
            new Paragraph(new Run(new Text("中国注册会计师：张三"))),
            new Paragraph(new Run(new Text("中国注册会计师：李四"))),
            new Paragraph(new Run(new Text("中国·成都"))),
            new Paragraph(new Run(new Text("二〇二六年十月二日"))),
            new SectionProperties());

        // 前置确认：该行确实被识别进落款区，且当前文本会被门禁按三级标题语义命中
        var signoff = new SignoffService().IdentifySignoffParagraphs(word, false);
        Assert.Contains(signoff, p => OpenXmlHelper.NormalizeText(OpenXmlHelper.ParagraphText(p)).StartsWith("1、", StringComparison.Ordinal));

        var gate = 排版后过门禁(word);

        Assert.DoesNotContain(gate.BlockingIssues, issue => issue.Code == "h3_separator_not_normalized");
        Assert.True(gate.Success, 拼接问题(gate));
    }

    [Fact]
    public void 纯封面文档_门禁不应做标题语义拦截()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocument(stream,
            new Paragraph(new Run(new Text("审计要求"))),
            new Paragraph(new Run(new Text("1.要点说明"))),
            new SectionProperties());

        var gate = 排版后过门禁(word);

        Assert.DoesNotContain(gate.BlockingIssues, issue => issue.Code == "h3_separator_not_normalized");
        Assert.True(gate.Success, 拼接问题(gate));
    }

    [Fact]
    public void 指向不存在部件的页眉死引用_排版后应清理()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocument(stream,
            正文填充段(),
            new SectionProperties(
                new HeaderReference { Id = "rIdGone", Type = HeaderFooterValues.Default }));

        var gate = 排版后过门禁(word);

        Assert.DoesNotContain(gate.BlockingIssues, issue => issue.Code == "header_ref_broken");
        Assert.DoesNotContain(gate.BlockingIssues, issue => issue.Code == "header_ref_missing_id");
        Assert.True(gate.Success, 拼接问题(gate));
    }

    [Fact]
    public void 缺少关系Id的页眉引用_排版后应清理()
    {
        using var stream = new MemoryStream();
        using var word = CreateDocument(stream,
            正文填充段(),
            new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default }));

        var gate = 排版后过门禁(word);

        Assert.DoesNotContain(gate.BlockingIssues, issue => issue.Code == "header_ref_missing_id");
        Assert.True(gate.Success, 拼接问题(gate));
    }

    // 必须用强类型 VML 构造：OpenXmlUnknownElement 树里的 w:t 不会实例化为 Text 节点，
    // 门禁的可见文本抽取会漏掉文字，导致测试假阴性；真实 Word 文档中文本框内段落是强类型的
    private static Paragraph CreateTextBoxParagraph(string textBoxText)
    {
        return new Paragraph(
            new Run(new Text("正文前缀")),
            new Run(new Picture(
                new V.Shape(
                    new V.TextBox(
                        new TextBoxContent(
                            new Paragraph(new Run(new Text(textBoxText))))))
                {
                    Id = "文本框1",
                    Style = "position:absolute;margin-left:10pt;margin-top:10pt;width:120pt;height:30pt"
                })));
    }
}
