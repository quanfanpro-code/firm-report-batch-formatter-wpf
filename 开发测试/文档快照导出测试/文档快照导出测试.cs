using System.Diagnostics;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Validation;
using DocxValidationTool;
using FirmFormatter.OpenXml.Core;
using Xunit;

namespace 文档快照导出测试;

public sealed class 文档快照导出测试
{
    [Fact]
    public void 生成矩阵时_应补出Task3新增坑点样本和说明()
    {
        var matrixRoot = Path.Combine(Path.GetTempPath(), $"矩阵_{Guid.NewGuid():N}");

        try
        {
            GenerateMatrixViaCli(matrixRoot);

            var inputDir = Path.Combine(matrixRoot, "Input");
            var expectedMd = Path.Combine(matrixRoot, "Expected", "矩阵说明.md");

            var expectedFiles = new[]
            {
                "标题坑点矩阵.docx",
                "表格坑点矩阵.docx",
                "页眉页脚坑点矩阵.docx",
                "落款坑点矩阵.docx",
                "非目标区域保护矩阵.docx",
                "文本框坑点矩阵.docx",
                "脚注尾注坑点矩阵.docx",
                "编号重启坑点矩阵.docx",
                "合并单元格坑点矩阵.docx",
                "落款复杂结构坑点矩阵.docx"
            };

            foreach (var fileName in expectedFiles)
            {
                Assert.True(File.Exists(Path.Combine(inputDir, fileName)), $"缺少新增矩阵样本：{fileName}");
            }

            var markdown = File.ReadAllText(expectedMd, Encoding.UTF8);
            Assert.Contains("标题坑点矩阵", markdown);
            Assert.Contains("表格坑点矩阵", markdown);
            Assert.Contains("页眉页脚坑点矩阵", markdown);
            Assert.Contains("落款坑点矩阵", markdown);
            Assert.Contains("非目标区域保护矩阵", markdown);
            Assert.Contains("文本框坑点矩阵", markdown);
            Assert.Contains("脚注尾注坑点矩阵", markdown);
            Assert.Contains("编号重启坑点矩阵", markdown);
            Assert.Contains("合并单元格坑点矩阵", markdown);
            Assert.Contains("落款复杂结构坑点矩阵", markdown);
        }
        finally
        {
            if (Directory.Exists(matrixRoot))
            {
                Directory.Delete(matrixRoot, true);
            }
        }
    }

    [Fact]
    public void 生成矩阵时_新增坑点样本应带上关键结构()
    {
        var matrixRoot = Path.Combine(Path.GetTempPath(), $"矩阵结构_{Guid.NewGuid():N}");

        try
        {
            GenerateMatrixViaCli(matrixRoot);

            AssertTitlePitfallMatrix(Path.Combine(matrixRoot, "Input", "标题坑点矩阵.docx"));
            AssertTablePitfallMatrix(Path.Combine(matrixRoot, "Input", "表格坑点矩阵.docx"));
            AssertHeaderFooterPitfallMatrix(Path.Combine(matrixRoot, "Input", "页眉页脚坑点矩阵.docx"));
            AssertSignoffPitfallMatrix(Path.Combine(matrixRoot, "Input", "落款坑点矩阵.docx"));
            AssertProtectedAreaMatrix(Path.Combine(matrixRoot, "Input", "非目标区域保护矩阵.docx"));
            AssertTextBoxPitfallMatrix(Path.Combine(matrixRoot, "Input", "文本框坑点矩阵.docx"));
            AssertFootnoteEndnotePitfallMatrix(Path.Combine(matrixRoot, "Input", "脚注尾注坑点矩阵.docx"));
            AssertNumberingRestartPitfallMatrix(Path.Combine(matrixRoot, "Input", "编号重启坑点矩阵.docx"));
            AssertMergedCellPitfallMatrix(Path.Combine(matrixRoot, "Input", "合并单元格坑点矩阵.docx"));
            AssertComplexSignoffPitfallMatrix(Path.Combine(matrixRoot, "Input", "落款复杂结构坑点矩阵.docx"));
        }
        finally
        {
            if (Directory.Exists(matrixRoot))
            {
                Directory.Delete(matrixRoot, true);
            }
        }
    }

    [Fact]
    public async Task 运行完整矩阵时_没有专用场景断言的合成样例不应导致整批失败()
    {
        var matrixRoot = Path.Combine(Path.GetTempPath(), $"矩阵运行_{Guid.NewGuid():N}");
        var repoRoot = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..", "..", ".."));
        var toolDll = Path.Combine(AppContext.BaseDirectory, "DocxValidation.dll");

        try
        {
            var startInfo = new ProcessStartInfo("dotnet")
            {
                WorkingDirectory = repoRoot,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            startInfo.ArgumentList.Add(toolDll);
            startInfo.ArgumentList.Add("run-matrix");
            startInfo.ArgumentList.Add(matrixRoot);

            using var process = Process.Start(startInfo);
            Assert.NotNull(process);
            var standardOutputTask = process!.StandardOutput.ReadToEndAsync();
            var standardErrorTask = process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();

            var standardOutput = await standardOutputTask;
            var standardError = await standardErrorTask;
            Assert.True(process.ExitCode == 0, $"运行矩阵失败。输出：{standardOutput} 错误：{standardError}");
            Assert.Contains("RESULT MatrixFailed=0", standardOutput);
        }
        finally
        {
            if (Directory.Exists(matrixRoot)) Directory.Delete(matrixRoot, true);
        }
    }

    [Fact]
    public void 导出快照时_应输出基础结构统计()
    {
        var docxPath = 矩阵测试资料.获取输出("标题矩阵_已排版.docx");

        var json = new DocumentSnapshotService().ExportAsJson(docxPath);
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;

        Assert.Equal(docxPath, root.GetProperty("文档路径").GetString());
        Assert.True(root.GetProperty("正文可见字数").GetInt32() > 0);
        Assert.True(root.GetProperty("段落数").GetInt32() > 0);
        Assert.True(root.GetProperty("表格数").GetInt32() >= 0);
        Assert.True(root.GetProperty("表格详情").GetArrayLength() >= 0);
        Assert.True(root.GetProperty("分节数").GetInt32() > 0);
        Assert.True(root.GetProperty("标题数").GetInt32() > 0);
        Assert.True(root.GetProperty("标题详情").GetArrayLength() > 0);
        Assert.True(root.GetProperty("分节详情").GetArrayLength() > 0);
        Assert.True(root.GetProperty("落款命中数").GetInt32() >= 0);
        Assert.True(root.TryGetProperty("复杂结构", out var complex));
        Assert.True(complex.TryGetProperty("文本框数", out _));
        Assert.True(complex.TryGetProperty("脚注数", out _));
        Assert.True(complex.TryGetProperty("尾注数", out _));
    }

    [Fact]
    public void 快照命令应写出Json文件()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..", "..", ".."));
        var docxPath = 矩阵测试资料.获取输出("标题矩阵_已排版.docx");
        var outputPath = Path.Combine(Path.GetTempPath(), $"快照_{Guid.NewGuid():N}.json");
        var toolDll = Path.Combine(AppContext.BaseDirectory, "DocxValidation.dll");

        try
        {
            var startInfo = new ProcessStartInfo("dotnet")
            {
                WorkingDirectory = repoRoot,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            startInfo.ArgumentList.Add(toolDll);
            startInfo.ArgumentList.Add("export-snapshot");
            startInfo.ArgumentList.Add(docxPath);
            startInfo.ArgumentList.Add(outputPath);

            using var process = Process.Start(startInfo);
            Assert.NotNull(process);
            process!.WaitForExit();

            var standardOutput = process.StandardOutput.ReadToEnd();
            var standardError = process.StandardError.ReadToEnd();

            Assert.True(process.ExitCode == 0, $"命令失败。输出：{standardOutput} 错误：{standardError}");
            Assert.True(File.Exists(outputPath), "快照文件没有生成");

            using var document = JsonDocument.Parse(File.ReadAllText(outputPath));
            Assert.True(document.RootElement.GetProperty("标题数").GetInt32() > 0);
            Assert.True(document.RootElement.GetProperty("分节详情").GetArrayLength() > 0);
        }
        finally
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    [Fact]
    public void 表格服务_应清掉原表残留样式并重建为README目标样式()
    {
        using var stream = new MemoryStream();
        using var word = CreateDirtyTableDocument(stream);

        var service = new TableService();
        service.Apply(word, hasCover: false);
        word.MainDocumentPart!.Document!.Save();

        var table = word.MainDocumentPart.Document.Body!.Elements<Table>().Single();
        var tableProps = table.GetFirstChild<TableProperties>();
        Assert.NotNull(tableProps);
        Assert.Null(tableProps!.GetFirstChild<TableStyle>());
        Assert.Null(tableProps.GetFirstChild<TableLook>());
        Assert.Null(tableProps.GetFirstChild<TableCellSpacing>());

        var tableBorders = tableProps.GetFirstChild<TableBorders>();
        Assert.NotNull(tableBorders);
        Assert.Equal(BorderValues.Single, tableBorders!.TopBorder?.Val?.Value);
        Assert.Equal(BorderValues.Single, tableBorders.BottomBorder?.Val?.Value);
        Assert.Equal(BorderValues.Dotted, tableBorders.InsideHorizontalBorder?.Val?.Value);
        Assert.Equal(BorderValues.Dotted, tableBorders.InsideVerticalBorder?.Val?.Value);
        Assert.Equal(BorderValues.Nil, tableBorders.LeftBorder?.Val?.Value);
        Assert.Equal(BorderValues.Nil, tableBorders.RightBorder?.Val?.Value);

        var rows = table.Elements<TableRow>().ToList();
        Assert.Equal(3, rows.Count);
        Assert.All(rows, row =>
        {
            Assert.Null(row.GetFirstChild<TableRowProperties>()?.GetFirstChild<TableCellSpacing>());
        });

        var firstRowCells = rows[0].Elements<TableCell>().ToList();
        Assert.All(firstRowCells, cell =>
        {
            var borders = cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<TableCellBorders>();
            Assert.NotNull(borders);
            Assert.Equal(BorderValues.Single, borders!.TopBorder?.Val?.Value);
            Assert.Equal(BorderValues.Dotted, borders.BottomBorder?.Val?.Value);
        });

        var middleRowCells = rows[1].Elements<TableCell>().ToList();
        Assert.All(middleRowCells, cell =>
        {
            var borders = cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<TableCellBorders>();
            Assert.NotNull(borders);
            Assert.Equal(BorderValues.Nil, borders!.TopBorder?.Val?.Value);
            Assert.Equal(BorderValues.Dotted, borders.BottomBorder?.Val?.Value);
        });

        var lastRowCells = rows[2].Elements<TableCell>().ToList();
        Assert.All(lastRowCells, cell =>
        {
            var borders = cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<TableCellBorders>();
            Assert.NotNull(borders);
            Assert.Equal(BorderValues.Nil, borders!.TopBorder?.Val?.Value);
            Assert.Equal(BorderValues.Single, borders.BottomBorder?.Val?.Value);
        });

        var validator = new OpenXmlValidator();
        var errors = validator.Validate(word).ToList();
        Assert.DoesNotContain(errors, error =>
            error.Path?.XPath?.Contains("/w:trPr", StringComparison.OrdinalIgnoreCase) == true &&
            error.Description.Contains("tblHeader", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void 表格服务_表头行已有对齐属性时_仍应生成合法的TableHeader顺序()
    {
        using var stream = new MemoryStream();
        using var word = CreateDirtyTableDocument(stream, withRowJustification: true);

        var service = new TableService();
        service.Apply(word, hasCover: false);
        word.MainDocumentPart!.Document!.Save();

        var table = word.MainDocumentPart.Document.Body!.Elements<Table>().Single();
        var firstRow = table.Elements<TableRow>().First();
        var trPr = firstRow.GetFirstChild<TableRowProperties>();
        Assert.NotNull(trPr);

        var childNames = trPr!.ChildElements.Select(x => x.LocalName).ToList();
        var headerIndex = childNames.IndexOf("tblHeader");
        var jcIndex = childNames.IndexOf("jc");
        Assert.True(headerIndex >= 0, "缺少 tblHeader");
        Assert.True(jcIndex >= 0, "缺少 jc");
        Assert.True(headerIndex < jcIndex, $"tblHeader 顺序错误：{string.Join(",", childNames)}");

        var validator = new OpenXmlValidator();
        var errors = validator.Validate(word).ToList();
        Assert.DoesNotContain(errors, error =>
            error.Path?.XPath?.Contains("/w:trPr", StringComparison.OrdinalIgnoreCase) == true &&
            error.Description.Contains("tblHeader", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void 表格服务_表头行已有后置变更节点时_TableHeader必须插在其前面()
    {
        using var stream = new MemoryStream();
        using var word = CreateDirtyTableDocument(stream, withRowJustification: false);
        var firstRow = word.MainDocumentPart!.Document!.Body!.Elements<Table>()
            .Single().Elements<TableRow>().First();
        var trPr = firstRow.GetFirstChild<TableRowProperties>() ?? firstRow.PrependChild(new TableRowProperties());
        trPr.AppendChild(new TableRowPropertiesChange { Id = "1" });

        new TableService().Apply(word, hasCover: false);

        var childNames = trPr.ChildElements.Select(x => x.LocalName).ToList();
        Assert.True(childNames.IndexOf("tblHeader") < childNames.IndexOf("trPrChange"),
            $"tblHeader 顺序错误：{string.Join(",", childNames)}");
    }

    [Fact]
    public void 表格服务_隐藏行属性必须保留在TableHeader之前且通过结构校验()
    {
        using var stream = new MemoryStream();
        using var word = CreateDirtyTableDocument(stream, withRowJustification: false);
        var firstRow = word.MainDocumentPart!.Document!.Body!.Elements<Table>()
            .Single().Elements<TableRow>().First();
        var trPr = firstRow.GetFirstChild<TableRowProperties>() ?? firstRow.PrependChild(new TableRowProperties());
        trPr.PrependChild(new Hidden());

        new TableService().Apply(word, hasCover: false);

        var childNames = trPr.ChildElements.Select(x => x.LocalName).ToList();
        Assert.True(childNames.IndexOf("hidden") < childNames.IndexOf("tblHeader"),
            $"tblHeader 顺序错误：{string.Join(",", childNames)}");
        AssertTableRowPropertiesValid(word);
    }

    [Fact]
    public void 表格服务_行前宽度属性必须位于TableHeader之后且通过结构校验()
    {
        using var stream = new MemoryStream();
        using var word = CreateDirtyTableDocument(stream, withRowJustification: false);
        var firstRow = word.MainDocumentPart!.Document!.Body!.Elements<Table>()
            .Single().Elements<TableRow>().First();
        var trPr = firstRow.GetFirstChild<TableRowProperties>() ?? firstRow.PrependChild(new TableRowProperties());
        trPr.RemoveAllChildren<TableCellSpacing>();
        trPr.AppendChild(new WidthBeforeTableRow { Width = "120", Type = TableWidthUnitValues.Dxa });

        new TableService().Apply(word, hasCover: false);

        var childNames = trPr.ChildElements.Select(x => x.LocalName).ToList();
        Assert.True(childNames.IndexOf("tblHeader") < childNames.IndexOf("wBefore"),
            $"tblHeader 顺序错误：{string.Join(",", childNames)}");
        AssertTableRowPropertiesValid(word);
    }

    private static void AssertTableRowPropertiesValid(WordprocessingDocument word)
    {
        var errors = new OpenXmlValidator().Validate(word)
            .Where(error => error.Path?.XPath?.Contains("/w:trPr", StringComparison.OrdinalIgnoreCase) == true)
            .ToList();
        Assert.True(errors.Count == 0,
            string.Join(Environment.NewLine, errors.Select(error => error.Description)));
    }

    private static void AssertTitlePitfallMatrix(string path)
    {
        using var word = WordprocessingDocument.Open(path, false);
        var paragraphs = word.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();

        var splitRunHeading = paragraphs.FirstOrDefault(p => GetVisibleText(p).Contains("拆分运行块一级标题", StringComparison.Ordinal));
        Assert.NotNull(splitRunHeading);
        Assert.True(splitRunHeading!.Elements<Run>().Count() >= 2, "标题坑点矩阵缺少拆分 run 标题");

        var styleChainHeading = paragraphs.FirstOrDefault(p => GetVisibleText(p).Contains("样式链挂编号但文本不带前缀", StringComparison.Ordinal));
        Assert.NotNull(styleChainHeading);
        Assert.Equal("Heading1", styleChainHeading!.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
    }

    private static void AssertTablePitfallMatrix(string path)
    {
        using var word = WordprocessingDocument.Open(path, false);
        var table = word.MainDocumentPart!.Document!.Body!.Elements<Table>().FirstOrDefault();
        Assert.NotNull(table);

        var rows = table!.Elements<TableRow>().ToList();
        Assert.True(rows.Count >= 4, "表格坑点矩阵行数不足");

        var complexCell = rows[1].Elements<TableCell>().ElementAt(1);
        Assert.True(complexCell.Elements<Paragraph>().Count() >= 2, "表格坑点矩阵缺少多段复杂单元格");

        var bookmarkCell = rows[2].Elements<TableCell>().ElementAt(1);
        Assert.True(bookmarkCell.Descendants<BookmarkStart>().Any(), "表格坑点矩阵缺少附着结构");

        var splitRunCell = rows[3].Elements<TableCell>().ElementAt(1);
        Assert.True(splitRunCell.Descendants<Run>().Count() >= 2, "表格坑点矩阵缺少拆分数字 run");
    }

    private static void AssertHeaderFooterPitfallMatrix(string path)
    {
        using var word = WordprocessingDocument.Open(path, false);
        var body = word.MainDocumentPart!.Document!.Body!;
        var sections = body.Descendants<SectionProperties>().ToList();

        Assert.True(sections.Count >= 2, "页眉页脚坑点矩阵分节不足");
        Assert.True(sections.Any(s => s.GetFirstChild<PageSize>()?.Orient?.Value == PageOrientationValues.Landscape), "页眉页脚坑点矩阵缺少横页分节");
        Assert.True(word.MainDocumentPart.HeaderParts.Count() >= 1, "页眉页脚坑点矩阵缺少页眉部件");
        Assert.True(word.MainDocumentPart.FooterParts.Count() >= 1, "页眉页脚坑点矩阵缺少页脚部件");

        var hasFieldCode = word.MainDocumentPart.FooterParts
            .SelectMany(p => p.RootElement?.Descendants<FieldCode>() ?? [])
            .Any();
        Assert.True(hasFieldCode, "页眉页脚坑点矩阵缺少域代码页脚");
    }

    private static void AssertSignoffPitfallMatrix(string path)
    {
        using var word = WordprocessingDocument.Open(path, false);
        var body = word.MainDocumentPart!.Document!.Body!;
        var paragraphs = body.Elements<Paragraph>().ToList();

        var cpaParagraph = paragraphs.FirstOrDefault(p => GetVisibleText(p).Contains("中国注册会计师", StringComparison.Ordinal));
        Assert.NotNull(cpaParagraph);
        Assert.True(cpaParagraph!.Descendants<TabChar>().Any(), "落款坑点矩阵缺少旧 tab");

        var specialLine = paragraphs.FirstOrDefault(p => GetVisibleText(p).Contains("特殊普通合伙", StringComparison.Ordinal));
        Assert.NotNull(specialLine);
        Assert.True(specialLine!.Descendants<Bold>().Any(), "落款坑点矩阵缺少旧粗体污染");
    }

    private static void AssertProtectedAreaMatrix(string path)
    {
        using var word = WordprocessingDocument.Open(path, false);
        var body = word.MainDocumentPart!.Document!.Body!;
        var table = body.Elements<Table>().FirstOrDefault();
        Assert.NotNull(table);

        var cells = table!.Descendants<TableCell>().ToList();
        Assert.True(cells.Any(c => GetVisibleText(c).Contains("00123", StringComparison.Ordinal)), "非目标区域保护矩阵缺少前导零文本");
        Assert.True(cells.Any(c => GetVisibleText(c).Contains("中国注册会计师", StringComparison.Ordinal)), "非目标区域保护矩阵缺少不应被正文表格逻辑误伤的文本");
    }

    private static void AssertTextBoxPitfallMatrix(string path)
    {
        using var word = WordprocessingDocument.Open(path, false);
        var body = word.MainDocumentPart!.Document!.Body!;
        Assert.Contains(body.Descendants<Paragraph>(), p =>
            OpenXmlHelper.提取文本框文本(p).Contains("文本框里的话", StringComparison.Ordinal));
    }

    private static void AssertFootnoteEndnotePitfallMatrix(string path)
    {
        using var word = WordprocessingDocument.Open(path, false);
        Assert.True(word.MainDocumentPart!.FootnotesPart?.Footnotes?.Descendants<Footnote>().Any(f => GetVisibleText(f).Contains("脚注内容", StringComparison.Ordinal)) == true, "脚注尾注坑点矩阵缺少脚注内容");
        Assert.True(word.MainDocumentPart.EndnotesPart?.Endnotes?.Descendants<Endnote>().Any(e => GetVisibleText(e).Contains("尾注内容", StringComparison.Ordinal)) == true, "脚注尾注坑点矩阵缺少尾注内容");
    }

    private static void AssertNumberingRestartPitfallMatrix(string path)
    {
        using var word = WordprocessingDocument.Open(path, false);
        var overrides = word.MainDocumentPart!.NumberingDefinitionsPart?.Numbering?.Descendants<LevelOverride>().ToList() ?? [];
        Assert.Contains(overrides, item => item.GetFirstChild<StartOverrideNumberingValue>()?.Val?.Value == 1);
    }

    private static void AssertMergedCellPitfallMatrix(string path)
    {
        using var word = WordprocessingDocument.Open(path, false);
        var body = word.MainDocumentPart!.Document!.Body!;
        Assert.True(body.Descendants<GridSpan>().Any(span => (span.Val?.Value ?? 1) > 1), "合并单元格坑点矩阵缺少横向合并");
        Assert.True(body.Descendants<VerticalMerge>().Any(), "合并单元格坑点矩阵缺少纵向合并");
    }

    private static void AssertComplexSignoffPitfallMatrix(string path)
    {
        using var word = WordprocessingDocument.Open(path, false);
        var body = word.MainDocumentPart!.Document!.Body!;
        Assert.Contains(body.Descendants<Paragraph>(), p => GetVisibleText(p).Contains("四川华信", StringComparison.Ordinal) && p.Descendants<BookmarkStart>().Any());
        Assert.Contains(body.Descendants<Paragraph>(), p => GetVisibleText(p).Contains("中国·成都", StringComparison.Ordinal) && p.Descendants<Hyperlink>().Any());
    }

    private static void GenerateMatrixViaCli(string matrixRoot)
    {
        Directory.CreateDirectory(matrixRoot);
        var repoRoot = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..", "..", ".."));
        var toolDll = Path.Combine(AppContext.BaseDirectory, "DocxValidation.dll");

        var startInfo = new ProcessStartInfo("dotnet")
        {
            WorkingDirectory = repoRoot,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.ArgumentList.Add(toolDll);
        startInfo.ArgumentList.Add("generate-matrix");
        startInfo.ArgumentList.Add(matrixRoot);

        using var process = Process.Start(startInfo);
        Assert.NotNull(process);
        process!.WaitForExit();

        var standardOutput = process.StandardOutput.ReadToEnd();
        var standardError = process.StandardError.ReadToEnd();
        Assert.True(process.ExitCode == 0, $"生成矩阵失败。输出：{standardOutput} 错误：{standardError}");
    }

    private static string GetVisibleText(OpenXmlElement element)
    {
        var builder = new StringBuilder();
        foreach (var text in element.Descendants<Text>())
        {
            builder.Append(text.Text);
        }

        return builder.ToString();
    }

    private static WordprocessingDocument CreateDirtyTableDocument(MemoryStream stream, bool withRowJustification = false)
    {
        var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var mainPart = word.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        var table = new Table(
            new TableProperties(
                new TableStyle { Val = "TableGrid" },
                new TableWidth { Width = "3333", Type = TableWidthUnitValues.Pct },
                new TableCellSpacing { Width = "15", Type = TableWidthUnitValues.Dxa },
                new TableLook { Val = "04A0", FirstRow = true, FirstColumn = true, NoVerticalBand = true },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Double, Size = 12, Color = "FF0000" },
                    new LeftBorder { Val = BorderValues.Single, Size = 8, Color = "FF0000" },
                    new BottomBorder { Val = BorderValues.Double, Size = 12, Color = "FF0000" },
                    new RightBorder { Val = BorderValues.Single, Size = 8, Color = "FF0000" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 8, Color = "FF0000" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 8, Color = "FF0000" })),
            new TableGrid(new GridColumn(), new GridColumn()));

        table.Append(
            CreateDirtyRow("项目", "金额", withRowJustification),
            CreateDirtyRow("营业收入（万元）", "3149.60", withRowJustification),
            CreateDirtyRow("净利润（万元）", "-548.20", withRowJustification));

        body.Append(table);
        mainPart.Document.Save();
        return word;
    }

    private static TableRow CreateDirtyRow(string leftText, string rightText, bool withRowJustification)
    {
        var rowProperties = new TableRowProperties(new TableCellSpacing { Width = "15", Type = TableWidthUnitValues.Dxa });
        if (withRowJustification)
        {
            rowProperties.Append(new Justification { Val = JustificationValues.Center });
        }

        return new TableRow(
            rowProperties,
            CreateDirtyCell(leftText, isFirstColumn: true),
            CreateDirtyCell(rightText, isFirstColumn: false));
    }

    private static TableCell CreateDirtyCell(string text, bool isFirstColumn)
    {
        return new TableCell(
            new TableCellProperties(
                new TableCellBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 6, Color = "000000" },
                    new LeftBorder { Val = isFirstColumn ? BorderValues.Nil : BorderValues.Single, Size = 6, Color = "000000" },
                    new BottomBorder { Val = BorderValues.Single, Size = 6, Color = "000000" },
                    new RightBorder { Val = BorderValues.Single, Size = 6, Color = "000000" })),
            new Paragraph(new Run(new Text(text))));
    }


    [Fact]
    public void 导出快照时_应输出分节方向页边距和页眉页脚引用()
    {
        var docxPath = 矩阵测试资料.获取输出("页眉页脚坑点矩阵_已排版.docx");

        var json = new DocumentSnapshotService().ExportAsJson(docxPath);
        using var document = JsonDocument.Parse(json);
        var sections = document.RootElement.GetProperty("分节详情");

        Assert.True(sections.GetArrayLength() >= 2);
        Assert.Contains(sections.EnumerateArray(), s => s.GetProperty("方向").GetString() == "横向");
        Assert.Contains(sections.EnumerateArray(), s => s.GetProperty("页眉引用数").GetInt32() >= 1);
        Assert.Contains(sections.EnumerateArray(), s => s.GetProperty("页脚引用数").GetInt32() >= 1);
        Assert.True(sections.EnumerateArray().All(s => s.TryGetProperty("页边距", out _)));
    }

    [Fact]
    public void 导出快照时_应输出落款命中与标题详情()
    {
        var docxPath = 矩阵测试资料.获取输出("落款坑点矩阵_已排版.docx");

        var json = new DocumentSnapshotService().ExportAsJson(docxPath);
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;

        Assert.True(root.GetProperty("落款命中数").GetInt32() >= 2);
        Assert.Contains(root.GetProperty("分节详情").EnumerateArray(), s => s.GetProperty("页脚引用数").GetInt32() >= 1);
    }

    [Fact]
    public void 对比相同快照时_不应产生差异()
    {
        var docxPath = 矩阵测试资料.获取输出("标题矩阵_已排版.docx");

        var snapshot = new DocumentSnapshotService().ExportAsJson(docxPath);
        var report = new DocumentDiffAuditService().CompareJson(snapshot, snapshot);

        Assert.Empty(report.差异列表);
    }

    [Fact]
    public void 对比不同快照时_应指出分节或标题差异()
    {
        var leftPath = 矩阵测试资料.获取输出("标题矩阵_已排版.docx");
        var rightPath = 矩阵测试资料.获取输出("页眉页脚坑点矩阵_已排版.docx");

        var left = new DocumentSnapshotService().ExportAsJson(leftPath);
        var right = new DocumentSnapshotService().ExportAsJson(rightPath);
        var report = new DocumentDiffAuditService().CompareJson(left, right);

        Assert.NotEmpty(report.差异列表);
        Assert.Contains(report.差异列表, item =>
            item.路径.Contains("分节", StringComparison.Ordinal) ||
            item.路径.Contains("标题", StringComparison.Ordinal));
        Assert.DoesNotContain(report.差异列表, item => item.路径.Contains("文档路径", StringComparison.Ordinal));
        Assert.True(report.区域差异数.Count > 0);
        Assert.True(report.区域差异数.ContainsKey("标题") || report.区域差异数.ContainsKey("分节"));
    }

    [Fact]
    public void 快照对比命令应写出差异Json文件()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..", "..", ".."));
        var leftPath = 矩阵测试资料.获取输出("标题矩阵_已排版.docx");
        var rightPath = 矩阵测试资料.获取输出("页眉页脚坑点矩阵_已排版.docx");
        var outputPath = Path.Combine(Path.GetTempPath(), $"差异_{Guid.NewGuid():N}.json");
        var toolDll = Path.Combine(AppContext.BaseDirectory, "DocxValidation.dll");

        try
        {
            var startInfo = new ProcessStartInfo("dotnet")
            {
                WorkingDirectory = repoRoot,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            startInfo.ArgumentList.Add(toolDll);
            startInfo.ArgumentList.Add("compare-snapshot");
            startInfo.ArgumentList.Add(leftPath);
            startInfo.ArgumentList.Add(rightPath);
            startInfo.ArgumentList.Add(outputPath);

            using var process = Process.Start(startInfo);
            Assert.NotNull(process);
            process!.WaitForExit();

            var standardOutput = process.StandardOutput.ReadToEnd();
            var standardError = process.StandardError.ReadToEnd();

            Assert.True(process.ExitCode == 0, $"命令失败。输出：{standardOutput} 错误：{standardError}");
            Assert.True(File.Exists(outputPath), "差异文件没有生成");

            using var document = JsonDocument.Parse(File.ReadAllText(outputPath));
            Assert.True(document.RootElement.GetProperty("差异列表").GetArrayLength() > 0);
        }
        finally
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }
}
