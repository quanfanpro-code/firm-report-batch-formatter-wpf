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

    [Fact]
    public void 只有序号列豁免数字格式化()
    {
        var table = new Table(
            new TableProperties(),
            new TableGrid(new GridColumn(), new GridColumn()),
            new TableRow(
                new TableCell(new Paragraph(new Run(new Text("金额")))),
                new TableCell(new Paragraph(new Run(new Text("序号"))))),
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
