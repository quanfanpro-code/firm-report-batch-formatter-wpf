using System.Diagnostics;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FirmFormatter.OpenXml.Contracts;
using FirmFormatter.OpenXml.Core;
using Xunit;
using V = DocumentFormat.OpenXml.Vml;

namespace 文档快照导出测试;

public sealed class 文件与结构独立复核测试
{
    private static Paragraph 段(string text) => new(new Run(new Text(text)));
    private static string 新目录()
    {
        var path = Path.Combine(Path.GetTempPath(), "文件结构复核", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }

    private static string 正常文件(string directory)
    {
        var path = Path.Combine(directory, "输入.docx");
        using var word = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        word.AddMainDocumentPart().Document = new Document(new Body(段(new string('正', 260)), new SectionProperties()));
        return path;
    }

    private static (int Code, string Output) 命令行(params string[] args)
    {
        var info = new ProcessStartInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "dotnet", "dotnet.exe"))
        { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false, CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8 };
        info.ArgumentList.Add(typeof(DocxValidationTool.CoverModeArgumentParser).Assembly.Location);
        foreach (var arg in args) info.ArgumentList.Add(arg);
        using var process = Process.Start(info)!;
        var output = process.StandardOutput.ReadToEndAsync();
        var error = process.StandardError.ReadToEndAsync();
        Assert.True(process.WaitForExit(30000), "命令行验证超时");
        return (process.ExitCode, output.GetAwaiter().GetResult() + error.GetAwaiter().GetResult());
    }

    [Theory]
    [InlineData("export-snapshot")]
    [InlineData("compare-snapshot")]
    public void 快照命令不得覆盖输入文件或已有结果(string command)
    {
        var directory = 新目录();
        var input = 正常文件(directory);
        var original = File.ReadAllBytes(input);
        var result = command == "export-snapshot" ? 命令行(command, input, input) : 命令行(command, input, input, input);
        Assert.NotEqual(0, result.Code);
        Assert.Equal(original, File.ReadAllBytes(input));
        var existing = Path.Combine(directory, "已有结果.json");
        File.WriteAllText(existing, "已有记录");
        result = command == "export-snapshot" ? 命令行(command, input, existing) : 命令行(command, input, input, existing);
        Assert.NotEqual(0, result.Code);
        Assert.Equal("已有记录", File.ReadAllText(existing));
    }

    [Fact]
    public void 独立验收命令必须执行业务门禁而不只是场景标题匹配()
    {
        var directory = 新目录();
        var input = 正常文件(directory);
        var output = Path.Combine(directory, "输出.docx");
        Assert.True(new DocumentPipeline(_ => { }).Run(new RequestContract { InputPath = input, OutputPath = output }).Success);
        // 页眉页脚场景以外的所有场景也必须检查页面业务规则。
        using (var word = WordprocessingDocument.Open(output, true))
            word.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single().GetFirstChild<PageMargin>()!.Left = 999U;
        var result = 命令行("verify-doc", output, "真实样本");
        Assert.NotEqual(0, result.Code);
        Assert.Contains("page_margin_top_left_right", result.Output);
    }

    [Fact]
    public void 仅有坏引用时只读门禁应返回问题而非直接崩溃()
    {
        using var stream = new MemoryStream();
        using var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        word.AddMainDocumentPart().Document = new Document(new Body(段("正文"),
            new SectionProperties(new HeaderReference { Id = "missing", Type = HeaderFooterValues.Default })));
        var xml = word.MainDocumentPart!.Document!.OuterXml;
        var result = new ValidationService().Validate(word, isPureCoverDocument: true, throwOnFailure: false);
        Assert.Contains(result.Issues, issue => issue.Code == "header_ref_broken");
        Assert.Equal(xml, word.MainDocumentPart!.Document!.OuterXml);
    }

    [Fact]
    public void 样式自动编号标题应出现在只读快照里()
    {
        var input = 正常文件(新目录());
        using (var word = WordprocessingDocument.Open(input, true))
        {
            var main = word.MainDocumentPart!;
            main.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(
                new AbstractNum(new Level(new NumberingFormat { Val = NumberFormatValues.ChineseCounting }, new LevelText { Val = "%1、" }) { LevelIndex = 0 }) { AbstractNumberId = 1 },
                new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new Style(new StyleParagraphProperties(
                new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 }))) { StyleId = "自定义章节", Type = StyleValues.Paragraph });
            main.Document!.Body!.PrependChild(new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "自定义章节" }), new Run(new Text("审计意见"))));
        }
        using var json = JsonDocument.Parse(new DocumentSnapshotService().ExportAsJson(input));
        Assert.Equal(1, json.RootElement.GetProperty("标题数").GetInt32());
    }

    [Fact]
    public void 数字前的特殊符号不得在表格改写时变成错误金额()
    {
        using var stream = new MemoryStream();
        using var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var cell = new TableCell(new Paragraph(new Run(new NoBreakHyphen(), new Text("1234"))));
        word.AddMainDocumentPart().Document = new Document(new Body(段(new string('正', 260)),
            new Table(new TableRow(new TableCell(段("项目")), new TableCell(段("金额"))),
                new TableRow(new TableCell(段("项目甲")), cell)), new SectionProperties()));
        new TableService().Apply(word, false);
        Assert.Equal("1234", cell.Descendants<Text>().Single().Text);
        Assert.Single(cell.Descendants<NoBreakHyphen>());
    }

    [Fact]
    public void 页眉规范化不得丢失特殊符号()
    {
        using var stream = new MemoryStream();
        using var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var main = word.AddMainDocumentPart();
        var header = main.AddNewPart<HeaderPart>();
        header.Header = new Header(new Paragraph(new Run(new Text("公司      报告"), new SymbolChar { Font = "Wingdings", Char = "F0FC" })));
        main.Document = new Document(new Body(段(new string('正', 260)), new SectionProperties(
            new HeaderReference { Id = main.GetIdOfPart(header), Type = HeaderFooterValues.Default })));
        new HeaderFooterService().Apply(word, false);
        Assert.Single(header.Header.Descendants<SymbolChar>());
    }

    [Fact]
    public void 跨分节共享页眉应分别按横竖版宽度定位()
    {
        using var stream = new MemoryStream();
        using var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var main = word.AddMainDocumentPart();
        var header = main.AddNewPart<HeaderPart>();
        header.Header = new Header(段("公司      报告"));
        var id = main.GetIdOfPart(header);
        main.Document = new Document(new Body(段(new string('正', 260)),
            new Paragraph(new ParagraphProperties(new SectionProperties(new HeaderReference { Id = id, Type = HeaderFooterValues.Default }))),
            段(new string('文', 260)), new SectionProperties(new HeaderReference { Id = id, Type = HeaderFooterValues.Default }, new PageSize { Orient = PageOrientationValues.Landscape })));
        new HeaderFooterService().Apply(word, false);
        var sections = OpenXmlHelper.收集分节(main.Document.Body);
        var first = (HeaderPart)main.GetPartById(sections[0].Elements<HeaderReference>().Single().Id!);
        var second = (HeaderPart)main.GetPartById(sections[1].Elements<HeaderReference>().Single().Id!);
        Assert.Equal(9639, first.Header!.Descendants<TabStop>().Single().Position?.Value);
        Assert.Equal(14571, second.Header!.Descendants<TabStop>().Single().Position?.Value);
    }

    [Fact]
    public void 删除签字图片必须被前后门禁拦截()
    {
        using var stream = new MemoryStream();
        using var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var image = new Picture(new V.Shape { Id = "签章" });
        word.AddMainDocumentPart().Document = new Document(new Body(段(new string('正', 260)), new Paragraph(new Run(image)), new SectionProperties()));
        var request = new RequestContract();
        var context = new FirmDocumentClassifier().BuildContext(word, request);
        new HeaderFooterService().Apply(word, false);
        Assert.True(new GateCheckService().Run(word, request, context).Success);
        image.Remove();
        Assert.False(new GateCheckService().Run(word, request, context).Success);
    }

    [Fact]
    public void 表格文字字号被改小必须拦截()
    {
        using var stream = new MemoryStream();
        using var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var table = new Table(new TableRow(new TableCell(段("金额"))), new TableRow(new TableCell(段("1234"))));
        word.AddMainDocumentPart().Document = new Document(new Body(段(new string('正', 260)), table, new SectionProperties()));
        var request = new RequestContract();
        var context = new FirmDocumentClassifier().BuildContext(word, request);
        new HeaderFooterService().Apply(word, false);
        new TableService().Apply(word, false);
        Assert.True(new GateCheckService().Run(word, request, context).Success);
        table.Descendants<Run>().Last().RunProperties!.FontSize!.Val = "18";
        Assert.False(new GateCheckService().Run(word, request, context).Success);
    }

    [Theory]
    [InlineData("缺少结束标记")]
    [InlineData("结束标记在前")]
    public void 页码域标记损坏必须拦截(string kind)
    {
        var input = 正常文件(新目录());
        using var word = WordprocessingDocument.Open(input, true);
        var request = new RequestContract();
        var context = new FirmDocumentClassifier().BuildContext(word, request);
        new HeaderFooterService().Apply(word, false);
        Assert.True(new GateCheckService().Run(word, request, context).Success);
        var p = word.MainDocumentPart!.FooterParts.Single().Footer!.Elements<Paragraph>().Single();
        var end = p.Descendants<FieldChar>().Single(f => f.FieldCharType?.Value == FieldCharValues.End);
        end.Remove();
        if (kind == "结束标记在前") p.InsertAfter(new Run(end), p.ParagraphProperties!);
        Assert.False(new GateCheckService().Run(word, request, context).Success);
    }

    [Theory]
    [InlineData("简单页码域")]
    [InlineData("首页不同明确关闭")]
    [InlineData("偶数页不同明确关闭")]
    public void 合法页脚变体必须放行(string kind)
    {
        var input = 正常文件(新目录());
        using var word = WordprocessingDocument.Open(input, true);
        var request = new RequestContract();
        var context = new FirmDocumentClassifier().BuildContext(word, request);
        new HeaderFooterService().Apply(word, false);
        if (kind == "简单页码域")
            word.MainDocumentPart!.FooterParts.Single().Footer = new Footer(new Paragraph(new SimpleField(new Run(new Text("1"))) { Instruction = " PAGE " }));
        if (kind == "首页不同明确关闭")
            word.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single().AppendChild(new TitlePage { Val = false });
        if (kind == "偶数页不同明确关闭")
            word.MainDocumentPart!.AddNewPart<DocumentSettingsPart>().Settings = new Settings(new EvenAndOddHeaders { Val = false });
        Assert.True(new GateCheckService().Run(word, request, context).Success);
    }

    [Fact]
    public void 仅承载文本框的空段落不能被文本框内编号误判成标题()
    {
        using var stream = new MemoryStream();
        using var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var inner = 段("1.备注说明");
        var outer = new Paragraph(new Run(new Picture(new V.Shape(new V.TextBox(new TextBoxContent(inner))))));
        word.AddMainDocumentPart().Document = new Document(new Body(段(new string('正', 260)), outer, new SectionProperties()));
        var request = new RequestContract();
        var context = new FirmDocumentClassifier().BuildContext(word, request);
        new HeaderFooterService().Apply(word, false);
        new ParagraphService().Apply(word, false, []);
        var result = new GateCheckService().Run(word, request, context);
        Assert.True(result.Success, string.Join("；", result.BlockingIssues.Select(i => i.Message)));
        Assert.Equal("1.备注说明", inner.InnerText);
        Assert.Null(outer.ParagraphProperties);
    }
}
