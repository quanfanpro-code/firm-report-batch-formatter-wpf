using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FirmFormatter.OpenXml.Contracts;
using FirmFormatter.OpenXml.Core;
using Xunit;

namespace 文档快照导出测试;

public sealed class 安全回归测试
{
    [Fact]
    public void 主窗口必须能够创建()
    {
        Exception? failure = null;
        var thread = new Thread(() =>
        {
            try
            {
                var window = new 事务所出报告批量排版WPF版.MainWindow();
                window.Close();
            }
            catch (Exception ex)
            {
                failure = ex;
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        Assert.Null(failure);
    }

    [Fact]
    public void 非Docx输入必须被拒绝()
    {
        var directory = 创建临时目录();
        var input = Path.Combine(directory, "输入.txt");
        File.WriteAllText(input, "不是 Word 文档");

        var error = new RequestContract
        {
            InputPath = input,
            OutputPath = Path.Combine(directory, "输出.docx")
        }.校验();

        Assert.Contains("docx", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void 批量处理必须跳过临时输出和失败件()
    {
        Assert.True(事务所出报告批量排版WPF版.输出文件命名规则.是已排版文件(
            $".报告.正在排版_{Guid.NewGuid():N}.docx"));
        Assert.True(事务所出报告批量排版WPF版.输出文件命名规则.是已排版文件(
            "报告_已排版_排版失败_20260809_160000.docx"));
        Assert.True(事务所出报告批量排版WPF版.输出文件命名规则.是已排版文件(
            "~$正在编辑的报告.docx"));
    }

    [Fact]
    public void OpenXml兼容性差异只警告不阻断()
    {
        using var stream = new MemoryStream();
        using var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var main = word.AddMainDocumentPart();
        main.Document = new Document(new Body(
            new Paragraph(new Run(new Text("有效正文"))),
            new SectionProperties(
                new PageSize { Width = 11906U, Height = 16838U },
                new PageMargin
                {
                    Top = 1417,
                    Bottom = 1701,
                    Left = 1417U,
                    Right = 850U,
                    Header = 850U,
                    Footer = 567U
                },
                new EvenAndOddHeaders())));
        new HeaderFooterService().Apply(word, hasCover: false);

        var report = new ValidationService().Validate(word, throwOnFailure: false);

        Assert.True(report.Success);
        Assert.Contains(report.Warnings, issue => issue.Code == "openxml_validator");
    }

    [Fact]
    public void 失败输出必须保留且不得占用正式输出名()
    {
        var directory = 创建临时目录();
        var input = Path.Combine(directory, "输入.docx");
        var output = Path.Combine(directory, "输入_已排版.docx");

        using (var word = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
        {
            var main = word.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new SectionProperties(),
                new Paragraph(new Run(new Text(new string('正', 260))))));
            main.Document.Save();
        }

        var result = new DocumentPipeline(_ => { }).Run(new RequestContract
        {
            InputPath = input,
            OutputPath = output
        });

        Assert.False(result.Success);
        Assert.False(File.Exists(output));
        Assert.Contains("排版失败", result.OutputPath, StringComparison.Ordinal);
        Assert.True(File.Exists(result.OutputPath), "失败件没有保留下来");
    }

    [Fact]
    public void 三级标题归一不得删除超链接()
    {
        var input = 创建正常文档(body => body.Append(
            new Paragraph(
                new Hyperlink(new Run(new Text("1."))) { Anchor = "正文书签" },
                new Run(new Text("三级标题")))));

        var result = 运行流水线(input);

        Assert.True(result.Success, result.Message);
        using var output = WordprocessingDocument.Open(result.OutputPath, false);
        var heading = output.MainDocumentPart!.Document!.Body!.Elements<Paragraph>()
            .Single(p => p.InnerText.Contains("三级标题", StringComparison.Ordinal));
        Assert.StartsWith("1．", heading.InnerText, StringComparison.Ordinal);
        Assert.True(heading.Descendants<Hyperlink>().Any(), "三级标题中的超链接被删除");
    }

    [Fact]
    public void 注册会计师复杂签字行不得删除超链接()
    {
        var input = 创建正常文档(body => body.Append(
            new Paragraph(new Run(new Text("四川华信(集团)会计师事务所"))),
            new Paragraph(new Run(new Text("（特殊普通合伙）"))),
            new Paragraph(new Run(new Text("中国·成都"))),
            new Paragraph(new Hyperlink(new Run(new Text("中国注册会计师：张三"))) { Anchor = "签字书签" }),
            new Paragraph(new Run(new Text("中国注册会计师：李四"))),
            new Paragraph(new Run(new Text("二〇二六年十月二日")))));

        var result = 运行流水线(input);

        Assert.True(result.Success, result.Message);
        using var output = WordprocessingDocument.Open(result.OutputPath, false);
        var signature = output.MainDocumentPart!.Document!.Body!.Elements<Paragraph>()
            .Single(p => p.InnerText.Contains("张三", StringComparison.Ordinal));
        Assert.True(signature.Descendants<Hyperlink>().Any(), "签字行中的超链接被删除");
    }

    [Theory]
    [InlineData("序号")]
    [InlineData("编号")]
    [InlineData("项目编号")]
    [InlineData("代码")]
    [InlineData("号码")]
    [InlineData("年份")]
    [InlineData("年度")]
    [InlineData("数量")]
    [InlineData("人数")]
    [InlineData("户数")]
    [InlineData("件数")]
    [InlineData("台数")]
    [InlineData("月份")]
    [InlineData("季度")]
    [InlineData("日期")]
    [InlineData("账龄")]
    [InlineData("期数")]
    [InlineData("页码")]
    [InlineData("账号")]
    [InlineData("银行账户")]
    [InlineData("识别号")]
    [InlineData("纳税人识别号")]
    [InlineData("银行账号")]
    public void 标识期间和数量类列必须豁免数字格式化(string header)
    {
        var table = new Table(
            new TableProperties(),
            new TableGrid(new GridColumn(), new GridColumn()),
            new TableRow(
                new TableCell(new Paragraph(new Run(new Text("金额")))),
                new TableCell(new Paragraph(new Run(new Text(header))))),
            new TableRow(
                new TableCell(new Paragraph(new Run(new Text("1234")))),
                new TableCell(new Paragraph(new Run(new Text("001"))))));
        var input = 创建正常文档(body => body.Append(table));

        var result = 运行流水线(input);

        Assert.True(result.Success, result.Message);
        using var output = WordprocessingDocument.Open(result.OutputPath, false);
        var cells = output.MainDocumentPart!.Document!.Body!.Descendants<TableRow>()
            .ElementAt(1).Elements<TableCell>().ToList();
        Assert.Equal("1,234.00", cells[0].InnerText);
        Assert.Equal("001", cells[1].InnerText);
    }

    [Fact]
    public void 封面正文同节时必须保留页码并处理正文表格()
    {
        var coverTable = new Table(
            new TableRow(new TableCell(new Paragraph(new Run(new Text("会计师事务所 地址 电话 传真"))))));
        var businessTable = new Table(
            new TableRow(new TableCell(new Paragraph(new Run(new Text("项目")))), new TableCell(new Paragraph(new Run(new Text("金额"))))),
            new TableRow(new TableCell(new Paragraph(new Run(new Text("收入")))), new TableCell(new Paragraph(new Run(new Text("1234"))))));
        var input = 创建正常文档(body => body.Append(coverTable, businessTable));

        var result = 运行流水线(input);

        Assert.True(result.Success, result.Message);
        using var output = WordprocessingDocument.Open(result.OutputPath, false);
        var body = output.MainDocumentPart!.Document!.Body!;
        Assert.NotEmpty(body.Descendants<FooterReference>());
        Assert.Equal("1,234.00", body.Elements<Table>().ElementAt(1).Descendants<TableCell>().Last().InnerText);
    }

    [Fact]
    public void 封面正文同节时必须识别正文落款()
    {
        var input = 创建正常文档(body => body.Append(
            new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("会计师事务所 地址 电话 传真")))))),
            new Paragraph(new Run(new Text("四川华信(集团)会计师事务所"))),
            new Paragraph(new Run(new Text("（特殊普通合伙）"))),
            new Paragraph(new Run(new Text("中国·成都"))),
            new Paragraph(new Run(new Text("中国注册会计师：张三"))),
            new Paragraph(new Run(new Text("中国注册会计师：李四"))),
            new Paragraph(new Run(new Text("二〇二六年十月二日")))));
        using var word = WordprocessingDocument.Open(input, false);

        var signoff = new SignoffService().IdentifySignoffParagraphs(word, hasCover: true);

        Assert.NotEmpty(signoff);
        Assert.Contains(signoff, paragraph => paragraph.InnerText.Contains("张三", StringComparison.Ordinal));
    }

    [Fact]
    public void 表格内自动编号必须保留()
    {
        var directory = 创建临时目录();
        var input = Path.Combine(directory, "输入.docx");
        using (var word = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
        {
            var main = word.AddMainDocumentPart();
            var numberedParagraph = new Paragraph(
                new ParagraphProperties(new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = 1 })),
                new Run(new Text("第一项")));
            var table = new Table(
                new TableRow(new TableCell(new Paragraph(new Run(new Text("事项"))))),
                new TableRow(new TableCell(numberedParagraph)));
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("审计报告"))),
                table,
                new Paragraph(new Run(new Text(new string('正', 260)))),
                new SectionProperties()));
            var numbering = main.AddNewPart<NumberingDefinitionsPart>();
            numbering.Numbering = new Numbering(
                new AbstractNum(new Level(
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%1." }) { LevelIndex = 0 }) { AbstractNumberId = 1 },
                new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
            main.Document.Save();
            numbering.Numbering.Save();
        }

        var result = 运行流水线(input);

        Assert.True(result.Success, result.Message);
        using var output = WordprocessingDocument.Open(result.OutputPath, false);
        var numbered = output.MainDocumentPart!.Document!.Body!.Descendants<Table>()
            .SelectMany(table => table.Descendants<Paragraph>()).Single(p => p.InnerText == "第一项");
        Assert.NotNull(numbered.ParagraphProperties?.GetFirstChild<NumberingProperties>());
    }

    [Fact]
    public void 奇偶页不同的文档必须同时生成偶数页页码()
    {
        using var stream = new MemoryStream();
        using var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var main = word.AddMainDocumentPart();
        main.Document = new Document(new Body(new Paragraph(new Run(new Text(new string('正', 260)))), new SectionProperties()));
        var settings = main.AddNewPart<DocumentSettingsPart>();
        settings.Settings = new Settings(new EvenAndOddHeaders());

        new HeaderFooterService().Apply(word, hasCover: false);

        var references = main.Document.Body!.Descendants<FooterReference>().ToList();
        Assert.Contains(references, item => item.Type?.Value == HeaderFooterValues.Default);
        Assert.Contains(references, item => item.Type?.Value == HeaderFooterValues.Even);
    }

    [Fact]
    public void 多行表头中的年度不得按金额格式化()
    {
        var table = new Table(
            new TableRow(new TableCell(new TableCellProperties(new GridSpan { Val = 2 }), new Paragraph(new Run(new Text("金额情况"))))),
            new TableRow(new TableCell(new Paragraph(new Run(new Text("项目")))), new TableCell(new Paragraph(new Run(new Text("2023"))))),
            new TableRow(new TableCell(new Paragraph(new Run(new Text("收入")))), new TableCell(new Paragraph(new Run(new Text("1234"))))));
        var input = 创建正常文档(body => body.Append(table));

        var result = 运行流水线(input);

        Assert.True(result.Success, result.Message);
        using var output = WordprocessingDocument.Open(result.OutputPath, false);
        var rows = output.MainDocumentPart!.Document!.Body!.Descendants<TableRow>().ToList();
        Assert.Equal("2023", rows[1].Elements<TableCell>().Last().InnerText);
        Assert.Equal("1,234.00", rows[2].Elements<TableCell>().Last().InnerText);
    }

    [Fact]
    public void 数据行含合并单元格时仍应按逻辑列豁免账号()
    {
        var table = new Table(
            new TableRow(
                new TableCell(new Paragraph(new Run(new Text("项目一")))),
                new TableCell(new Paragraph(new Run(new Text("项目二")))),
                new TableCell(new Paragraph(new Run(new Text("银行账号"))))),
            new TableRow(
                new TableCell(new TableCellProperties(new GridSpan { Val = 2 }), new Paragraph(new Run(new Text("开户信息")))),
                new TableCell(new Paragraph(new Run(new Text("001234"))))));
        var input = 创建正常文档(body => body.Append(table));

        var result = 运行流水线(input);

        Assert.True(result.Success, result.Message);
        using var output = WordprocessingDocument.Open(result.OutputPath, false);
        Assert.Equal("001234", output.MainDocumentPart!.Document!.Body!.Descendants<TableRow>()
            .ElementAt(1).Elements<TableCell>().Last().InnerText);
    }

    [Fact]
    public void 账户余额仍应按金额格式化()
    {
        var table = new Table(
            new TableRow(new TableCell(new Paragraph(new Run(new Text("项目")))), new TableCell(new Paragraph(new Run(new Text("账户余额"))))),
            new TableRow(new TableCell(new Paragraph(new Run(new Text("银行存款")))), new TableCell(new Paragraph(new Run(new Text("1234"))))));
        var input = 创建正常文档(body => body.Append(table));

        var result = 运行流水线(input);

        Assert.True(result.Success, result.Message);
        using var output = WordprocessingDocument.Open(result.OutputPath, false);
        Assert.Equal("1,234.00", output.MainDocumentPart!.Document!.Body!.Descendants<TableRow>()
            .ElementAt(1).Elements<TableCell>().Last().InnerText);
    }

    [Fact]
    public void 多层账户表头必须以末层账号和余额分别判断()
    {
        var table = new Table(
            new TableRow(new TableCell(new TableCellProperties(new GridSpan { Val = 2 }), new Paragraph(new Run(new Text("银行账户"))))),
            new TableRow(
                new TableCell(new Paragraph(new Run(new Text("账号")))),
                new TableCell(new Paragraph(new Run(new Text("余额"))))),
            new TableRow(
                new TableCell(new Paragraph(new Run(new Text("001234")))),
                new TableCell(new Paragraph(new Run(new Text("1234"))))));
        var input = 创建正常文档(body => body.Append(table));

        var result = 运行流水线(input);

        Assert.True(result.Success, result.Message);
        using var output = WordprocessingDocument.Open(result.OutputPath, false);
        var cells = output.MainDocumentPart!.Document!.Body!.Descendants<TableRow>()
            .ElementAt(2).Elements<TableCell>().ToList();
        Assert.Equal("001234", cells[0].InnerText);
        Assert.Equal("1,234.00", cells[1].InnerText);
    }

    [Fact]
    public void 表头省略前置网格列时账号豁免坐标不得错位()
    {
        var table = new Table(
            new TableRow(
                new TableRowProperties(new GridBefore { Val = 1 }),
                new TableCell(new Paragraph(new Run(new Text("账号")))),
                new TableCell(new Paragraph(new Run(new Text("余额"))))),
            new TableRow(
                new TableCell(new Paragraph(new Run(new Text("项目")))),
                new TableCell(new Paragraph(new Run(new Text("001234")))),
                new TableCell(new Paragraph(new Run(new Text("1234"))))));
        var input = 创建正常文档(body => body.Append(table));

        var result = 运行流水线(input);

        Assert.True(result.Success, result.Message);
        using var output = WordprocessingDocument.Open(result.OutputPath, false);
        var cells = output.MainDocumentPart!.Document!.Body!.Descendants<TableRow>()
            .ElementAt(1).Elements<TableCell>().ToList();
        Assert.Equal("001234", cells[1].InnerText);
        Assert.Equal("1,234.00", cells[2].InnerText);
    }

    [Fact]
    public void 合并数据格覆盖账号列时不得按金额改写()
    {
        var table = new Table(
            new TableRow(
                new TableCell(new Paragraph(new Run(new Text("项目")))),
                new TableCell(new Paragraph(new Run(new Text("账号")))),
                new TableCell(new Paragraph(new Run(new Text("金额"))))),
            new TableRow(
                new TableCell(new TableCellProperties(new GridSpan { Val = 2 }), new Paragraph(new Run(new Text("001234")))),
                new TableCell(new Paragraph(new Run(new Text("5678"))))));
        var input = 创建正常文档(body => body.Append(table));

        var result = 运行流水线(input);

        Assert.True(result.Success, result.Message);
        using var output = WordprocessingDocument.Open(result.OutputPath, false);
        var cells = output.MainDocumentPart!.Document!.Body!.Descendants<TableRow>()
            .ElementAt(1).Elements<TableCell>().ToList();
        Assert.Equal("001234", cells[0].InnerText);
        Assert.Equal("5,678.00", cells[1].InnerText);
    }

    [Fact]
    public void 数字单元格含隐藏文字或脚注引用时不得重写复杂内容()
    {
        var hiddenText = new Text("内部标记");
        var footnoteRun = new Run(
            new RunProperties(new RunStyle { Val = "FootnoteReference" }),
            new FootnoteReference { Id = 1 });
        var table = new Table(
            new TableRow(new TableCell(new Paragraph(new Run(new Text("项目")))), new TableCell(new Paragraph(new Run(new Text("金额"))))),
            new TableRow(
                new TableCell(new Paragraph(new Run(new Text("收入")))),
                new TableCell(new Paragraph(
                    new Run(new Text("1234")),
                    new Run(new RunProperties(new Vanish()), hiddenText),
                    footnoteRun))));
        var input = 创建正常文档(body => body.Append(table));

        var result = 运行流水线(input);

        Assert.True(result.Success, result.Message);
        using var output = WordprocessingDocument.Open(result.OutputPath, false);
        var amountCell = output.MainDocumentPart!.Document!.Body!.Descendants<TableRow>()
            .ElementAt(1).Elements<TableCell>().Last();
        Assert.Contains(amountCell.Descendants<Text>(), text => text.Text == "内部标记");
        Assert.True(amountCell.Descendants<FootnoteReference>().Any());
        Assert.Equal("FootnoteReference", amountCell.Descendants<FootnoteReference>().Single()
            .Ancestors<Run>().Single().RunProperties?.RunStyle?.Val?.Value);
    }

    [Fact]
    public void 普通正文中的脚注引用样式必须保留()
    {
        using var stream = new MemoryStream();
        using var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var main = word.AddMainDocumentPart();
        var referenceRun = new Run(
            new RunProperties(new RunStyle { Val = "FootnoteReference" }),
            new FootnoteReference { Id = 1 });
        main.Document = new Document(new Body(
            new Paragraph(new Run(new Text("普通正文")), referenceRun),
            new Paragraph(new Run(new Text(new string('正', 260)))),
            new SectionProperties()));

        new ParagraphService().Apply(word, hasCover: false, []);

        Assert.Equal("FootnoteReference", referenceRun.RunProperties?.RunStyle?.Val?.Value);
    }

    [Fact]
    public void 复杂签字行中的换行隐藏文字和脚注引用不得丢失()
    {
        using var stream = new MemoryStream();
        using var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var main = word.AddMainDocumentPart();
        var signature = new Paragraph(
            new Run(new Text("中国注册会计师：张三")),
            new Run(new Break()),
            new Run(new RunProperties(new Vanish()), new Text("内部标记")),
            new Run(new FootnoteReference { Id = 1 }));
        var date = new Paragraph(new Run(new Text("二〇二六年十月二日")));
        main.Document = new Document(new Body(signature, date, new SectionProperties()));

        new SignoffService().Apply(word, [signature, date]);

        Assert.True(signature.Descendants<Break>().Any());
        Assert.True(signature.Descendants<Vanish>().Any());
        Assert.True(signature.Descendants<FootnoteReference>().Any());
    }

    [Fact]
    public void 超过两位小数的数字改写必须写入抽查日志()
    {
        var logs = new List<LogEventContract>();
        var table = new Table(
            new TableRow(new TableCell(new Paragraph(new Run(new Text("项目")))), new TableCell(new Paragraph(new Run(new Text("金额"))))),
            new TableRow(new TableCell(new Paragraph(new Run(new Text("收入")))), new TableCell(new Paragraph(new Run(new Text("1.234"))))));
        var input = 创建正常文档(body => body.Append(table));
        var output = Path.Combine(Path.GetDirectoryName(input)!, "输出.docx");

        var result = new DocumentPipeline(logs.Add).Run(new RequestContract { InputPath = input, OutputPath = output });

        Assert.True(result.Success, result.Message);
        Assert.Contains(logs, log => log.Code == "numeric_rounding" && log.Message.Contains("1", StringComparison.Ordinal));
    }

    [Fact]
    public void 未识别到落款区必须写入人工复核警告()
    {
        var logs = new List<LogEventContract>();
        var input = 创建正常文档(_ => { });
        var output = Path.Combine(Path.GetDirectoryName(input)!, "输出.docx");

        var result = new DocumentPipeline(logs.Add).Run(new RequestContract { InputPath = input, OutputPath = output });

        Assert.True(result.Success, result.Message);
        Assert.Contains(logs, log => log.Code == "signoff_not_found" && log.Type == "warning");
    }

    [Fact]
    public void 目录条目不得按正文标题重排()
    {
        using var stream = new MemoryStream();
        using var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var main = word.AddMainDocumentPart();
        var toc = new Paragraph(
            new ParagraphProperties(
                new ParagraphStyleId { Val = "TOC1" },
                new Tabs(new TabStop { Val = TabStopValues.Right, Position = 9000 }),
                new Justification { Val = JustificationValues.Right }),
            new Run(new Text("一、审计意见")),
            new Run(new TabChar()),
            new Run(new Text("1")));
        main.Document = new Document(new Body(
            toc,
            new Paragraph(new Run(new Text(new string('正', 260)))),
            new SectionProperties()));

        new ParagraphService().Apply(word, hasCover: false, []);

        Assert.Equal(JustificationValues.Right, toc.ParagraphProperties?.Justification?.Val?.Value);
        Assert.NotNull(toc.ParagraphProperties?.Tabs);
    }

    [Fact]
    public void 已取消的请求不得继续生成输出()
    {
        var input = 创建正常文档(_ => { });
        var output = Path.Combine(Path.GetDirectoryName(input)!, "输出.docx");
        var request = new RequestContract { InputPath = input, OutputPath = output };
        var cancellationProperty = typeof(RequestContract).GetProperty("CancellationToken");
        Assert.NotNull(cancellationProperty);
        cancellationProperty.SetValue(request, new CancellationToken(canceled: true));

        var result = new DocumentPipeline(_ => { }).Run(request);

        Assert.False(result.Success);
        Assert.Equal("cancelled", result.ErrorCode);
        Assert.False(File.Exists(output));
    }

    private static ResponseContract 运行流水线(string input)
    {
        var output = Path.Combine(Path.GetDirectoryName(input)!, "输出.docx");
        return new DocumentPipeline(_ => { }).Run(new RequestContract
        {
            InputPath = input,
            OutputPath = output
        });
    }

    private static string 创建正常文档(Action<Body> appendBusinessContent)
    {
        var directory = 创建临时目录();
        var input = Path.Combine(directory, "输入.docx");
        using var word = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document);
        var main = word.AddMainDocumentPart();
        var body = new Body();
        main.Document = new Document(body);
        body.Append(new Paragraph(new Run(new Text("审计报告"))));
        appendBusinessContent(body);
        body.Append(
            new Paragraph(new Run(new Text(new string('正', 260)))),
            new SectionProperties());
        main.Document.Save();
        return input;
    }

    private static string 创建临时目录()
    {
        var directory = Path.Combine(Path.GetTempPath(), "FirmFormatterSafetyTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return directory;
    }
}
