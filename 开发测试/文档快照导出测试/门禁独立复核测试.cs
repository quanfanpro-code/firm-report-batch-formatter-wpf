using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FirmFormatter.OpenXml.Contracts;
using FirmFormatter.OpenXml.Core;
using Xunit;
using V = DocumentFormat.OpenXml.Vml;

namespace 文档快照导出测试;

// 独立构造正常件和损坏件，断言对外结果，避免用排版函数计算预期值。
public sealed class 门禁独立复核测试
{
    private static Paragraph 段(string text) => new(new Run(new Text(text)));

    private static WordprocessingDocument 文档(MemoryStream stream, params OpenXmlElement[] children)
    {
        var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        word.AddMainDocumentPart().Document = new Document(new Body(
            children.Concat(new OpenXmlElement[] { 段(new string('正', 260)), new SectionProperties() })));
        return word;
    }

    private static FirmDocumentContext 排版(WordprocessingDocument word, bool? cover = null)
    {
        var request = new RequestContract { HasCoverOverride = cover, ScenarioName = "常规" };
        var context = new FirmDocumentClassifier().BuildContext(word, request);
        new HeaderFooterService().Apply(word, context.HasCover);
        var signoff = new SignoffService().IdentifySignoffParagraphs(word, context.HasCover);
        new ParagraphService().Apply(word, context.HasCover, signoff);
        new TableService().Apply(word, context.HasCover);
        new SignoffService().Apply(word, signoff);
        return context;
    }

    private static GateCheckResultContract 门禁(WordprocessingDocument word, FirmDocumentContext context)
        => new GateCheckService().Run(word, context.Request, context);

    private static void 应通过(GateCheckResultContract result)
        => Assert.True(result.Success, string.Join("；", result.BlockingIssues.Select(i => i.Code + ":" + i.Message)));

    [Fact]
    public void 多个坏页眉引用必须全部清理且保留后面的好页眉()
    {
        using var stream = new MemoryStream();
        using var word = 文档(stream);
        var main = word.MainDocumentPart!;
        var header = main.AddNewPart<HeaderPart>();
        header.Header = new Header(段("有效页眉"));
        var section = main.Document!.Body!.Elements<SectionProperties>().Single();
        section.Append(new HeaderReference { Type = HeaderFooterValues.First },
            new HeaderReference { Id = "不存在", Type = HeaderFooterValues.Even },
            new HeaderReference { Id = main.GetIdOfPart(header), Type = HeaderFooterValues.Default });
        var context = 排版(word);
        应通过(门禁(word, context));
        Assert.Single(section.Elements<HeaderReference>());
        Assert.Equal("21", header.Header.Descendants<Run>().Single().RunProperties?.FontSize?.Val?.Value);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void 页眉页脚指向样式部件必须报错且不得删除样式(bool headerReference)
    {
        using var stream = new MemoryStream();
        using var word = 文档(stream);
        var main = word.MainDocumentPart!;
        var styles = main.AddNewPart<StyleDefinitionsPart>();
        styles.Styles = new Styles(new Style { StyleId = "保留样式", Type = StyleValues.Paragraph });
        var section = main.Document!.Body!.Elements<SectionProperties>().Single();
        section.Append(headerReference
            ? new HeaderReference { Id = main.GetIdOfPart(styles), Type = HeaderFooterValues.Default }
            : new FooterReference { Id = main.GetIdOfPart(styles), Type = HeaderFooterValues.Default });
        var before = new ValidationService().Validate(word, throwOnFailure: false);
        Assert.Contains(before.Issues, i => i.Code == (headerReference ? "header_ref_wrong_type" : "footer_ref_wrong_type"));
        排版(word);
        Assert.NotNull(main.StyleDefinitionsPart);
        Assert.Contains(main.StyleDefinitionsPart!.Styles!.Elements<Style>(), s => s.StyleId?.Value == "保留样式");
    }

    [Fact]
    public void 封面表格数字不排版也不应被业务门禁误拦()
    {
        using var stream = new MemoryStream();
        var table = new Table(new TableRow(new TableCell(段("会计师事务所 地址 电话 传真"))),
            new TableRow(new TableCell(段("12345"))));
        using var word = 文档(stream, table, new Paragraph(new ParagraphProperties(new SectionProperties())), 段("一、审计意见"));
        var original = table.OuterXml;
        var context = 排版(word);
        Assert.True(context.HasCover);
        Assert.Equal(original, table.OuterXml);
        应通过(门禁(word, context));
    }

    [Fact]
    public void 删除全部分节后门禁不得空循环放行()
    {
        using var stream = new MemoryStream();
        using var word = 文档(stream);
        var context = 排版(word);
        word.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single().Remove();
        Assert.False(门禁(word, context).Success);
    }

    [Fact]
    public void 没有显式末节的合法输入也必须实际设置页边距和页码()
    {
        using var stream = new MemoryStream();
        using var word = 文档(stream);
        word.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single().Remove();
        var context = 排版(word);
        应通过(门禁(word, context));
        var section = Assert.Single(word.MainDocumentPart.Document!.Body.Elements<SectionProperties>());
        Assert.Equal(1417, section.GetFirstChild<PageMargin>()?.Top?.Value);
        Assert.NotEmpty(section.Elements<FooterReference>());
    }

    [Fact]
    public void 页脚只留固定数字而无页码域必须拦截()
    {
        using var stream = new MemoryStream();
        using var word = 文档(stream);
        var context = 排版(word);
        word.MainDocumentPart!.FooterParts.Single().Footer = new Footer(段("1"));
        Assert.False(门禁(word, context).Success);
    }

    [Fact]
    public void 首页不同却缺少首页页脚必须拦截()
    {
        using var stream = new MemoryStream();
        using var word = 文档(stream);
        var section = word.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();
        section.Append(new TitlePage());
        var context = 排版(word);
        section.Elements<FooterReference>().Single(f => f.Type?.Value == HeaderFooterValues.First).Remove();
        Assert.False(门禁(word, context).Success);
    }

    [Fact]
    public void 非A4页面必须拦截()
    {
        using var stream = new MemoryStream();
        using var word = 文档(stream);
        var context = 排版(word);
        word.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single().GetFirstChild<PageSize>()!.Width = 12240U;
        Assert.False(门禁(word, context).Success);
    }

    [Fact]
    public void 显式取消隐藏的文字仍是可见正文()
    {
        var p = new Paragraph(new Run(new RunProperties(new Vanish { Val = false }, new WebHidden { Val = false }), new Text("可见正文")));
        Assert.Equal("可见正文", OpenXmlHelper.ParagraphText(p));
    }

    [Fact]
    public void 三级标题前有隐藏文字仍应规范且不改隐藏内容()
    {
        using var stream = new MemoryStream();
        var hidden = new Text("内部标记");
        var heading = new Paragraph(new Run(new RunProperties(new Vanish()), hidden), new Run(new Text("1.基本情况")));
        using var word = 文档(stream, heading);
        var context = 排版(word);
        应通过(门禁(word, context));
        Assert.Equal("内部标记", hidden.Text);
        Assert.Equal("1．基本情况", OpenXmlHelper.ParagraphText(heading));
    }

    [Fact]
    public void 标准普通类型脚注被删除也必须拦截()
    {
        using var stream = new MemoryStream();
        using var word = 文档(stream);
        var part = word.MainDocumentPart!.AddNewPart<FootnotesPart>();
        part.Footnotes = new Footnotes(new Footnote(段("脚注原文")) { Id = 1, Type = FootnoteEndnoteValues.Normal });
        var context = 排版(word);
        part.Footnotes.Elements<Footnote>().Single().Remove();
        Assert.False(门禁(word, context).Success);
    }

    [Fact]
    public void 正文脚注引用被删除即使注释正文还在也必须拦截()
    {
        using var stream = new MemoryStream();
        var reference = new FootnoteReference { Id = 1 };
        using var word = 文档(stream, new Paragraph(new Run(new Text("有脚注"), reference)));
        var part = word.MainDocumentPart!.AddNewPart<FootnotesPart>();
        part.Footnotes = new Footnotes(new Footnote(段("脚注原文")) { Id = 1 });
        var context = 排版(word);
        reference.Remove();
        Assert.False(门禁(word, context).Success);
    }

    [Fact]
    public void 简单域丢失必须拦截()
    {
        using var stream = new MemoryStream();
        var field = new SimpleField(new Run(new Text("引用结果"))) { Instruction = " REF 原文书签 " };
        using var word = 文档(stream, new Paragraph(field));
        var context = 排版(word);
        field.Remove();
        Assert.False(门禁(word, context).Success);
    }

    [Fact]
    public void 两个同一宿主段落的文本框应分别计数()
    {
        using var stream = new MemoryStream();
        using var word = 文档(stream, new Paragraph(new Run(new Picture(
            new V.Shape(new V.TextBox(new TextBoxContent(段("甲")))),
            new V.Shape(new V.TextBox(new TextBoxContent(段("乙"))))))));
        Assert.Equal(2, 复杂结构观察服务.观察(word).文本框数);
    }

    [Fact]
    public void 文本框外层段落排版不得改到文本框内部字号()
    {
        using var stream = new MemoryStream();
        var inner = new Run(new RunProperties(new FontSize { Val = "18" }), new Text("备注"));
        using var word = 文档(stream, new Paragraph(new Run(new Text("正常正文")), new Run(new Picture(
            new V.Shape(new V.TextBox(new TextBoxContent(new Paragraph(inner))))))));
        var context = 排版(word);
        应通过(门禁(word, context));
        Assert.Equal("18", inner.RunProperties?.FontSize?.Val?.Value);
    }

    [Fact]
    public void 重复排版页眉必须保留右定位点()
    {
        using var stream = new MemoryStream();
        using var word = 文档(stream);
        var main = word.MainDocumentPart!;
        var header = main.AddNewPart<HeaderPart>();
        header.Header = new Header(段("公司名称        审计报告"));
        main.Document!.Body!.Elements<SectionProperties>().Single().Append(new HeaderReference { Id = main.GetIdOfPart(header), Type = HeaderFooterValues.Default });
        排版(word);
        排版(word);
        var p = header.Header.Elements<Paragraph>().Single();
        Assert.Single(p.Descendants<TabChar>());
        Assert.Equal(9639, p.ParagraphProperties?.Tabs?.Elements<TabStop>().Single().Position?.Value);
    }

    [Fact]
    public void 页眉中的隐藏文字不得因统一字体变成可见()
    {
        using var stream = new MemoryStream();
        using var word = 文档(stream);
        var main = word.MainDocumentPart!;
        var header = main.AddNewPart<HeaderPart>();
        header.Header = new Header(new Paragraph(new Run(new Text("公司名称")), new Run(new RunProperties(new Vanish()), new Text("内部备注"))));
        main.Document!.Body!.Elements<SectionProperties>().Single().Append(new HeaderReference { Id = main.GetIdOfPart(header), Type = HeaderFooterValues.Default });
        排版(word);
        Assert.Equal("公司名称", OpenXmlHelper.提取可见文本(header.Header));
    }

    [Fact]
    public void 真实样本验收不能写死某份报告的章节名称()
    {
        using var stream = new MemoryStream();
        using var word = 文档(stream, 段("一、专项审核意见"));
        var context = 排版(word);
        var request = new RequestContract { ScenarioName = "真实样本", RequireScenarioVerification = true };
        context = new FirmDocumentClassifier().BuildContext(word, request);
        应通过(门禁(word, context));
    }

    [Theory]
    [InlineData("字号")]
    [InlineData("缩进")]
    [InlineData("行距")]
    public void 标题格式损坏必须拦截(string kind)
    {
        using var stream = new MemoryStream();
        var heading = 段("一、审计意见");
        using var word = 文档(stream, heading);
        var context = 排版(word);
        if (kind == "字号") heading.Descendants<Run>().Single().RunProperties!.FontSize!.Val = "18";
        if (kind == "缩进") heading.ParagraphProperties!.Indentation!.FirstLineChars = 200;
        if (kind == "行距") heading.ParagraphProperties!.SpacingBetweenLines!.Line = "240";
        Assert.False(门禁(word, context).Success);
    }

    [Fact]
    public void 落款阶梯缩进损坏必须拦截()
    {
        using var stream = new MemoryStream();
        using var word = 文档(stream, 段("四川华信(集团)会计师事务所"), 段("（特殊普通合伙）"),
            段("中国·成都"), 段("中国注册会计师：张三"), 段("中国注册会计师：李四"), 段("二〇二六年九月七日"));
        var context = 排版(word);
        word.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ElementAt(1).ParagraphProperties!.Indentation!.FirstLineChars = 0;
        Assert.False(门禁(word, context).Success);
    }

    [Fact]
    public void 门禁通过日志时收到取消也不得发布正式输出()
    {
        var folder = Path.Combine(Path.GetTempPath(), "门禁独立复核", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        var input = Path.Combine(folder, "输入.docx");
        var output = Path.Combine(folder, "输出.docx");
        using (var word = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
            word.AddMainDocumentPart().Document = new Document(new Body(段(new string('正', 260)), new SectionProperties()));
        using var cts = new CancellationTokenSource();
        var result = new DocumentPipeline(e => { if (e.Code == "validation_done") cts.Cancel(); }).Run(new RequestContract
        { InputPath = input, OutputPath = output, CancellationToken = cts.Token });
        Assert.False(result.Success);
        Assert.Equal("cancelled", result.ErrorCode);
        Assert.False(File.Exists(output));
        Assert.True(File.Exists(input));
    }
}
