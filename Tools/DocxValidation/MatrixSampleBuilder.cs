using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxValidationTool;

public static class MatrixSampleBuilder
{
    public static void GenerateAll(string matrixRoot)
    {
        var inputDir = Path.Combine(matrixRoot, "Input");
        var outputDir = Path.Combine(matrixRoot, "Output");
        var expectedDir = Path.Combine(matrixRoot, "Expected");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);
        Directory.CreateDirectory(expectedDir);

        GenerateHeadingMatrix(Path.Combine(inputDir, "标题矩阵.docx"));
        GenerateBodyNumberingMatrix(Path.Combine(inputDir, "正文编号矩阵.docx"));
        GenerateTableMatrix(Path.Combine(inputDir, "表格复杂文本矩阵.docx"));
        GenerateCoverMatrix(Path.Combine(inputDir, "封面矩阵.docx"));
        GeneratePureCoverMatrix(Path.Combine(inputDir, "纯封面矩阵.docx"));
        GenerateSignoffMatrix(Path.Combine(inputDir, "落款矩阵.docx"));
        GenerateSharedNumberingMatrix(Path.Combine(inputDir, "共享编号模板防误伤矩阵.docx"));
        GenerateHeaderFooterMatrix(Path.Combine(inputDir, "页眉页脚分节矩阵.docx"));
        GenerateHeadingPitfallMatrix(Path.Combine(inputDir, "标题坑点矩阵.docx"));
        GenerateTablePitfallMatrix(Path.Combine(inputDir, "表格坑点矩阵.docx"));
        GenerateHeaderFooterPitfallMatrix(Path.Combine(inputDir, "页眉页脚坑点矩阵.docx"));
        GenerateSignoffPitfallMatrix(Path.Combine(inputDir, "落款坑点矩阵.docx"));
        GenerateProtectedAreaMatrix(Path.Combine(inputDir, "非目标区域保护矩阵.docx"));
        GenerateTextBoxPitfallMatrix(Path.Combine(inputDir, "文本框坑点矩阵.docx"));
        GenerateFootnoteEndnotePitfallMatrix(Path.Combine(inputDir, "脚注尾注坑点矩阵.docx"));
        GenerateNumberingRestartPitfallMatrix(Path.Combine(inputDir, "编号重启坑点矩阵.docx"));
        GenerateMergedCellPitfallMatrix(Path.Combine(inputDir, "合并单元格坑点矩阵.docx"));
        GenerateComplexSignoffPitfallMatrix(Path.Combine(inputDir, "落款复杂结构坑点矩阵.docx"));

        File.WriteAllText(
            Path.Combine(expectedDir, "矩阵说明.md"),
            BuildExpectedMarkdown(),
            Encoding.UTF8);
    }

    private static void GenerateHeadingMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        AddBaseStyles(stylesPart);

        var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new Numbering();
        AddHeadingStyleNumbering(numberingPart, stylesPart);

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("标题矩阵样本", styleId: "Title"));
        body.Append(Paragraph("一、文本前缀一级标题"));
        body.Append(Paragraph("（一）文本前缀二级标题"));
        body.Append(Paragraph("1.文本前缀三级标题"));
        body.Append(Paragraph("样式链一级标题", styleId: "Heading1"));
        body.Append(Paragraph("样式链二级标题", styleId: "Heading2"));
        body.Append(Paragraph("样式链三级标题", styleId: "Heading3"));
        body.Append(NumberedParagraph("段落自身一级标题", 41, 0));
        body.Append(NumberedParagraph("段落自身二级标题", 41, 1));
        body.Append(NumberedParagraph("段落自身三级标题", 41, 2));
        body.Append(OutlineParagraph("仅大纲一级标题", 0));
        body.Append(OutlineParagraph("仅大纲二级标题", 1));
        body.Append(OutlineParagraph("仅大纲三级标题", 2));
        AppendPaddingBody(body, "标题矩阵填充正文");
        AppendBodySectionProperties(body);
        mainPart.Document.Save();
    }

    private static void GenerateBodyNumberingMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        AddBaseStyles(stylesPart);

        var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new Numbering();
        AddBodyNumbering(numberingPart, stylesPart);

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("正文自动编号矩阵", styleId: "Title"));
        body.Append(Paragraph("以下内容不属于一级、二级、三级标题，但必须保留编号语义。"));
        body.Append(NumberedParagraph("这是正文式四级（直接编号）", 61, 0));
        body.Append(NumberedParagraph("这是正文式四级（第二项）", 61, 0));
        body.Append(Paragraph("这是正文式四级（样式链）", styleId: "BodyLevel4"));
        body.Append(Paragraph("这是正文式四级（样式链第二项）", styleId: "BodyLevel4"));
        body.Append(NumberedParagraph("这是普通说明列表（直接编号）", 62, 0));
        body.Append(NumberedParagraph("这是普通说明列表（第二项）", 62, 0));
        body.Append(Paragraph("这是普通说明列表（样式链）", styleId: "BodyList"));
        body.Append(Paragraph("这是普通说明列表（样式链第二项）", styleId: "BodyList"));
        AppendPaddingBody(body, "正文编号矩阵填充正文");
        AppendBodySectionProperties(body);
        mainPart.Document.Save();
    }

    private static void GenerateTableMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        AddBaseStyles(stylesPart);

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("表格复杂文本矩阵", styleId: "Title"));

        var table = new Table(
            new TableProperties(
                new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 6, Color = "000000" },
                    new BottomBorder { Val = BorderValues.Single, Size = 6, Color = "000000" },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" })),
            new TableGrid(new GridColumn(), new GridColumn(), new GridColumn()));

        table.Append(
            Row(Cell("项目"), Cell("期末数"), Cell("备注")),
            Row(Cell("00123"), Cell("1234"), Cell("首列正整数字符串")),
            Row(Cell("A100"), Cell("1", "234"), Cell("拆分运行块数字")),
            Row(Cell("B200"), ComplexCellWithTabsAndMultipleParagraphs(), Cell("复杂结构应保守处理")),
            Row(Cell("0"), Cell("12.5%"), Cell("零值与百分比")));

        body.Append(table);
        body.Append(new Paragraph());
        AppendPaddingBody(body, "表格矩阵填充正文");
        AppendBodySectionProperties(body);
        mainPart.Document.Save();
    }

    private static void GenerateCoverMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        AddBaseStyles(stylesPart);

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("封面矩阵样本", styleId: "Title"));
        body.Append(BuildCoverKeywordTable(splitRuns: true));
        body.Append(Paragraph("封面结束说明"));

        var coverSectionParagraph = Paragraph("封面与正文分节点");
        coverSectionParagraph.ParagraphProperties ??= new ParagraphProperties();
        coverSectionParagraph.ParagraphProperties.Append(BuildSectionProperties());
        body.Append(coverSectionParagraph);

        body.Append(Paragraph("财务报表附注", styleId: "Title"));
        body.Append(Paragraph("样式链一级标题", styleId: "Heading1"));
        AppendPaddingBody(body, "封面矩阵正文填充");
        AppendBodySectionProperties(body);
        mainPart.Document.Save();
    }

    private static void GeneratePureCoverMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("财务报表附注", styleId: null));
        body.Append(Paragraph("华辰示例私募基金管理有限公司"));
        body.Append(Paragraph("2025年度"));
        AppendBodySectionProperties(body);
        mainPart.Document.Save();
    }

    private static void GenerateSignoffMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        AddBaseStyles(stylesPart);

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("落款矩阵样本", styleId: "Title"));
        AppendPaddingBody(body, "落款前正文");

        body.Append(ParagraphFromRuns("四川华信", "(集团)", "会计师事务所"));
        body.Append(Paragraph("（特殊普通合伙）"));
        body.Append(ParagraphFromRuns("中国", "·", "成都"));
        body.Append(Paragraph("中国注册会计师：张三"));
        body.Append(Paragraph("中国注册会计师：李四"));
        body.Append(ParagraphFromRuns("二〇二六年", "X", "月", "X", "日"));
        AppendBodySectionProperties(body);
        mainPart.Document.Save();
    }

    private static void GenerateSharedNumberingMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        AddBaseStyles(stylesPart);

        var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new Numbering();
        AddSharedNumberingDefinition(numberingPart);

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("共享编号模板防误伤矩阵", styleId: "Title"));
        body.Append(NumberedParagraph("1.共享模板短标题", 81, 0));
        body.Append(NumberedParagraph("这是一个长度明显超过三十个字的长编号正文，用来锁死共享 numbering 模板时不能误把长正文清坏。", 81, 0));
        AppendPaddingBody(body, "共享编号模板矩阵正文");
        AppendBodySectionProperties(body);
        mainPart.Document.Save();
    }

    private static void GenerateHeaderFooterMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        AddBaseStyles(stylesPart);

        var sharedHeaderPart = mainPart.AddNewPart<HeaderPart>();
        sharedHeaderPart.Header = new Header(new Paragraph(new Run(new Text("共享页眉"))));
        sharedHeaderPart.Header.Save();
        var sharedHeaderId = mainPart.GetIdOfPart(sharedHeaderPart);

        var sharedFooterPart = mainPart.AddNewPart<FooterPart>();
        sharedFooterPart.Footer = new Footer(new Paragraph(new Run(new Text("共享页脚"))));
        sharedFooterPart.Footer.Save();
        var sharedFooterId = mainPart.GetIdOfPart(sharedFooterPart);

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("页眉页脚分节矩阵", styleId: "Title"));
        body.Append(BuildCoverKeywordTable(splitRuns: false));
        body.Append(Paragraph("封面结束说明"));

        var coverSectionParagraph = Paragraph("封面分节");
        coverSectionParagraph.ParagraphProperties ??= new ParagraphProperties();
        coverSectionParagraph.ParagraphProperties.Append(BuildSectionProperties(sharedHeaderId, sharedFooterId));
        body.Append(coverSectionParagraph);

        body.Append(Paragraph("财务报表附注", styleId: "Title"));
        body.Append(Paragraph("正文第一页", styleId: "Heading1"));
        AppendPaddingBody(body, "页眉页脚矩阵正文");
        body.Append(BuildBodySectionProperties(sharedHeaderId, sharedFooterId));
        mainPart.Document.Save();
    }

    private static void GenerateHeadingPitfallMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        AddBaseStyles(stylesPart);

        var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new Numbering();
        AddHeadingStyleNumbering(numberingPart, stylesPart);

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("标题坑点矩阵", styleId: "Title"));
        body.Append(ParagraphFromRuns("一、", "拆分运行块一级标题"));
        body.Append(ParagraphFromRuns("（一）", "拆分运行块二级标题"));
        body.Append(ParagraphFromRuns("1", "．", "拆分运行块三级标题"));
        body.Append(Paragraph("样式链挂编号但文本不带前缀", styleId: "Heading1"));
        body.Append(Paragraph("样式链挂编号但文本不带二级前缀", styleId: "Heading2"));
        body.Append(Paragraph("样式链挂编号但文本不带三级前缀", styleId: "Heading3"));
        AppendPaddingBody(body, "标题坑点矩阵填充正文");
        AppendBodySectionProperties(body);
        mainPart.Document.Save();
    }

    private static void GenerateTablePitfallMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        AddBaseStyles(stylesPart);

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("表格坑点矩阵", styleId: "Title"));

        var table = new Table(
            new TableProperties(
                new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 6, Color = "000000" },
                    new BottomBorder { Val = BorderValues.Single, Size = 6, Color = "000000" },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" })),
            new TableGrid(new GridColumn(), new GridColumn(), new GridColumn()));

        table.Append(
            Row(Cell("项目"), Cell("原始内容"), Cell("意图")),
            Row(Cell("多段复杂单元格"), ComplexCellWithTabsAndMultipleParagraphs(), Cell("不能被压扁")),
            Row(Cell("附着结构"), BookmarkCell("带书签的结构化文本", "表格书签"), Cell("不能误删附着结构")),
            Row(Cell("拆分数字"), Cell("1", "234"), Cell("仍需识别成同一个数值单元格")));

        body.Append(table);
        body.Append(new Paragraph());
        AppendPaddingBody(body, "表格坑点矩阵填充正文");
        AppendBodySectionProperties(body);
        mainPart.Document.Save();
    }

    private static void GenerateHeaderFooterPitfallMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        AddBaseStyles(stylesPart);

        var headerPart = mainPart.AddNewPart<HeaderPart>();
        headerPart.Header = new Header(
            new Paragraph(
                new Run(new Text("旧页眉污染")),
                new Run(new Text(" / ")),
                new Run(new Text("共享页眉样本"))));
        headerPart.Header.Save();
        var headerId = mainPart.GetIdOfPart(headerPart);

        var footerPart = mainPart.AddNewPart<FooterPart>();
        footerPart.Footer = new Footer(
            new Paragraph(
                new Run(new Text("第")),
                new Run(new FieldCode(" PAGE ")),
                new Run(new Text("页"))));
        footerPart.Footer.Save();
        var footerId = mainPart.GetIdOfPart(footerPart);

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("页眉页脚坑点矩阵", styleId: "Title"));
        body.Append(Paragraph("封面节保留共享部件引用，用于验证后续清理边界。"));

        var portraitSectionParagraph = Paragraph("竖页节结束");
        portraitSectionParagraph.ParagraphProperties ??= new ParagraphProperties();
        portraitSectionParagraph.ParagraphProperties.Append(BuildSectionProperties(headerId, footerId));
        body.Append(portraitSectionParagraph);

        body.Append(Paragraph("横页正文开始", styleId: "Heading1"));
        body.Append(Paragraph("这里用于制造横竖页混排和共享页眉页脚部件。"));

        var landscapeSectionParagraph = Paragraph("横页节结束");
        landscapeSectionParagraph.ParagraphProperties ??= new ParagraphProperties();
        landscapeSectionParagraph.ParagraphProperties.Append(BuildLandscapeSectionProperties(headerId, footerId));
        body.Append(landscapeSectionParagraph);

        body.Append(Paragraph("恢复竖页正文", styleId: "Heading1"));
        AppendPaddingBody(body, "页眉页脚坑点矩阵填充正文");
        body.Append(BuildBodySectionProperties(headerId, footerId));
        mainPart.Document.Save();
    }

    private static void GenerateSignoffPitfallMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        AddBaseStyles(stylesPart);

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("落款坑点矩阵", styleId: "Title"));
        AppendPaddingBody(body, "落款坑点矩阵前正文");
        body.Append(ParagraphFromRuns("四川华信", "    ", "(集团)", "会计师事务所"));
        body.Append(BoldParagraph("（特殊普通合伙）"));
        body.Append(ParagraphFromRuns("中国", "·", "成都"));
        body.Append(TabParagraph("中国注册会计师：张三", "签字位"));
        body.Append(TabParagraph("中国注册会计师：李四", "签字位"));
        body.Append(ParagraphFromRuns("二〇二六年", " ", "十", " ", "月", " ", "二", " ", "日"));
        AppendBodySectionProperties(body);
        mainPart.Document.Save();
    }

    private static void GenerateProtectedAreaMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        AddBaseStyles(stylesPart);

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("非目标区域保护矩阵", styleId: "Title"));
        body.Append(Paragraph("这个样本用于锁死前导零文本、说明性表格和落款关键字误伤问题。"));

        var table = new Table(
            new TableProperties(
                new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 6, Color = "000000" },
                    new BottomBorder { Val = BorderValues.Single, Size = 6, Color = "000000" },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" })),
            new TableGrid(new GridColumn(), new GridColumn()));

        table.Append(
            Row(Cell("编号"), Cell("说明")),
            Row(Cell("00123"), Cell("前导零文本必须原样保留")),
            Row(Cell("中国注册会计师：这里是表格说明"), Cell("这不是落款区，不能被落款逻辑误伤")),
            Row(Cell("1.", "这是共享编号模板里的长正文说明"), Cell("不能被误判成短标题")));

        body.Append(table);
        body.Append(new Paragraph());
        AppendPaddingBody(body, "非目标区域保护矩阵填充正文");
        AppendBodySectionProperties(body);
        mainPart.Document.Save();
    }

    private static void GenerateTextBoxPitfallMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        AddBaseStyles(stylesPart);

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("文本框坑点矩阵", styleId: "Title"));
        body.Append(TextBoxParagraph(mainPart, "正文前缀", "文本框里的话"));
        AppendPaddingBody(body, "文本框坑点矩阵填充正文");
        AppendBodySectionProperties(body);
        mainPart.Document.Save();
    }

    private static void GenerateFootnoteEndnotePitfallMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        AddBaseStyles(stylesPart);

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("脚注尾注坑点矩阵", styleId: "Title"));
        body.Append(new Paragraph(
            new ParagraphProperties(),
            new Run(new Text("正文脚注位置")),
            new Run(new FootnoteReference { Id = 1 })));
        body.Append(new Paragraph(
            new ParagraphProperties(),
            new Run(new Text("正文尾注位置")),
            new Run(new EndnoteReference { Id = 1 })));

        var footnotesPart = mainPart.AddNewPart<FootnotesPart>();
        footnotesPart.Footnotes = new Footnotes(
            new Footnote { Type = FootnoteEndnoteValues.Separator, Id = -1 },
            new Footnote(new Paragraph(new Run(new Text("脚注内容")))) { Id = 1 });
        footnotesPart.Footnotes.Save();

        var endnotesPart = mainPart.AddNewPart<EndnotesPart>();
        endnotesPart.Endnotes = new Endnotes(
            new Endnote { Type = FootnoteEndnoteValues.Separator, Id = -1 },
            new Endnote(new Paragraph(new Run(new Text("尾注内容")))) { Id = 1 });
        endnotesPart.Endnotes.Save();

        AppendPaddingBody(body, "脚注尾注坑点矩阵填充正文");
        AppendBodySectionProperties(body);
        mainPart.Document.Save();
    }

    private static void GenerateNumberingRestartPitfallMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        AddBaseStyles(stylesPart);

        var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new Numbering();
        AddRestartNumberingDefinition(numberingPart);

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("编号重启坑点矩阵", styleId: "Title"));
        body.Append(NumberedParagraph("第一组编号正文", 92, 0));
        body.Append(NumberedParagraph("第一组编号正文第二项", 92, 0));
        body.Append(NumberedParagraph("重启后的编号正文", 93, 0));
        AppendPaddingBody(body, "编号重启坑点矩阵填充正文");
        AppendBodySectionProperties(body);
        mainPart.Document.Save();
    }

    private static void GenerateMergedCellPitfallMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        AddBaseStyles(stylesPart);

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("合并单元格坑点矩阵", styleId: "Title"));
        body.Append(CreateMergedTable());
        AppendPaddingBody(body, "合并单元格坑点矩阵填充正文");
        AppendBodySectionProperties(body);
        mainPart.Document.Save();
    }

    private static void GenerateComplexSignoffPitfallMatrix(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        AddBaseStyles(stylesPart);

        var body = mainPart.Document.Body!;
        body.Append(Paragraph("落款复杂结构坑点矩阵", styleId: "Title"));
        AppendPaddingBody(body, "落款复杂结构坑点矩阵前正文");
        foreach (var paragraph in CreateComplexSignoffParagraphs())
        {
            body.Append(paragraph);
        }
        AppendBodySectionProperties(body);
        mainPart.Document.Save();
    }

    private static void AddBaseStyles(StyleDefinitionsPart stylesPart)
    {
        var styles = stylesPart.Styles!;
        styles.Append(
            new DocDefaults(
                new RunPropertiesDefault(
                    new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "宋体" },
                        new FontSize { Val = "24" },
                        new FontSizeComplexScript { Val = "24" })),
                new ParagraphPropertiesDefault(
                    new ParagraphPropertiesBaseStyle(
                        new SpacingBetweenLines { Line = "420", LineRule = LineSpacingRuleValues.Exact }))));

        styles.Append(Style("Normal", "正文", null, null, 24, false, null));
        styles.Append(Style("Title", "标题", "Normal", null, 32, true, null, JustificationValues.Center));
        styles.Append(Style("Heading1", "标题 1", "Normal", 0, 24, true, null));
        styles.Append(Style("Heading2", "标题 2", "Normal", 1, 24, true, null));
        styles.Append(Style("Heading3", "标题 3", "Normal", 2, 24, true, "480"));
        styles.Append(Style("BodyLevel4", "正文四级", "Normal", null, 24, false, "480"));
        styles.Append(Style("BodyList", "正文列表", "Normal", null, 24, false, null));
        styles.Save();
    }

    private static Style Style(
        string styleId,
        string name,
        string? basedOn,
        int? outlineLevel,
        int size,
        bool bold,
        string? firstLine,
        JustificationValues? justification = null)
    {
        var style = new Style { Type = StyleValues.Paragraph, StyleId = styleId };
        style.Append(new StyleName { Val = name });
        if (!string.IsNullOrWhiteSpace(basedOn))
        {
            style.Append(new BasedOn { Val = basedOn });
        }

        style.Append(new NextParagraphStyle { Val = "Normal" });
        style.Append(new UIPriority { Val = 9 });
        style.Append(new PrimaryStyle());

        var styleParagraphProperties = new StyleParagraphProperties(
            new SpacingBetweenLines { Line = "420", LineRule = LineSpacingRuleValues.Exact });
        if (!string.IsNullOrWhiteSpace(firstLine))
        {
            styleParagraphProperties.Indentation = new Indentation { FirstLine = firstLine };
        }
        styleParagraphProperties.Justification = new Justification { Val = justification ?? JustificationValues.Both };
        if (outlineLevel.HasValue)
        {
            styleParagraphProperties.OutlineLevel = new OutlineLevel { Val = outlineLevel.Value };
        }

        style.Append(styleParagraphProperties);

        var runProperties = new StyleRunProperties(
            new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "宋体" },
            new Bold { Val = bold },
            new BoldComplexScript { Val = bold },
            new FontSize { Val = size.ToString(CultureInfo.InvariantCulture) },
            new FontSizeComplexScript { Val = size.ToString(CultureInfo.InvariantCulture) });

        style.Append(runProperties);
        return style;
    }

    private static void AddHeadingStyleNumbering(NumberingDefinitionsPart numberingPart, StyleDefinitionsPart stylesPart)
    {
        numberingPart.Numbering!.Append(
            new AbstractNum(
                HeadingLevel(0, NumberFormatValues.ChineseCounting, "%1、"),
                HeadingLevel(1, NumberFormatValues.ChineseCountingThousand, "（%2）"),
                HeadingLevel(2, NumberFormatValues.Decimal, "%3．"))
            { AbstractNumberId = 40, MultiLevelType = new MultiLevelType { Val = MultiLevelValues.Multilevel } });

        numberingPart.Numbering.Append(new NumberingInstance(new AbstractNumId { Val = 40 }) { NumberID = 41 });
        BindStyleNumbering(stylesPart.Styles!, "Heading1", 41, 0);
        BindStyleNumbering(stylesPart.Styles!, "Heading2", 41, 1);
        BindStyleNumbering(stylesPart.Styles!, "Heading3", 41, 2);
        numberingPart.Numbering.Save();
    }

    private static void AddBodyNumbering(NumberingDefinitionsPart numberingPart, StyleDefinitionsPart stylesPart)
    {
        numberingPart.Numbering!.Append(
            new AbstractNum(
                HeadingLevel(0, NumberFormatValues.Decimal, "（%1）"))
            { AbstractNumberId = 60, MultiLevelType = new MultiLevelType { Val = MultiLevelValues.HybridMultilevel } });

        numberingPart.Numbering.Append(
            new AbstractNum(
                new Level(
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%1." },
                    new LevelJustification { Val = LevelJustificationValues.Left })
                { LevelIndex = 0 })
            { AbstractNumberId = 61, MultiLevelType = new MultiLevelType { Val = MultiLevelValues.HybridMultilevel } });

        numberingPart.Numbering.Append(new NumberingInstance(new AbstractNumId { Val = 60 }) { NumberID = 61 });
        numberingPart.Numbering.Append(new NumberingInstance(new AbstractNumId { Val = 61 }) { NumberID = 62 });
        BindStyleNumbering(stylesPart.Styles!, "BodyLevel4", 61, 0);
        BindStyleNumbering(stylesPart.Styles!, "BodyList", 62, 0);
        numberingPart.Numbering.Save();
    }

    private static void AddSharedNumberingDefinition(NumberingDefinitionsPart numberingPart)
    {
        numberingPart.Numbering!.Append(
            new AbstractNum(
                new Level(
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%1." },
                    new LevelJustification { Val = LevelJustificationValues.Left })
                { LevelIndex = 0 })
            { AbstractNumberId = 80, MultiLevelType = new MultiLevelType { Val = MultiLevelValues.HybridMultilevel } });

        numberingPart.Numbering.Append(new NumberingInstance(new AbstractNumId { Val = 80 }) { NumberID = 81 });
        numberingPart.Numbering.Save();
    }

    private static void AddRestartNumberingDefinition(NumberingDefinitionsPart numberingPart)
    {
        numberingPart.Numbering!.Append(
            new AbstractNum(
                new Level(
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%1." },
                    new LevelJustification { Val = LevelJustificationValues.Left })
                { LevelIndex = 0 })
            { AbstractNumberId = 90, MultiLevelType = new MultiLevelType { Val = MultiLevelValues.HybridMultilevel } });

        numberingPart.Numbering.Append(new NumberingInstance(new AbstractNumId { Val = 90 }) { NumberID = 92 });
        numberingPart.Numbering.Append(
            new NumberingInstance(
                new AbstractNumId { Val = 90 },
                new LevelOverride(
                    new StartOverrideNumberingValue { Val = 1 })
                { LevelIndex = 0 })
            { NumberID = 93 });
        numberingPart.Numbering.Save();
    }

    private static Level HeadingLevel(int level, NumberFormatValues format, string levelText)
    {
        return new Level(
            new NumberingFormat { Val = format },
            new LevelText { Val = levelText },
            new LevelJustification { Val = LevelJustificationValues.Left })
        { LevelIndex = level };
    }

    private static void BindStyleNumbering(Styles styles, string styleId, int numId, int ilvl)
    {
        var style = styles.Elements<Style>().First(s => s.StyleId?.Value == styleId);
        style.StyleParagraphProperties ??= new StyleParagraphProperties();
        style.StyleParagraphProperties.NumberingProperties?.Remove();
        var numberingProperties = new NumberingProperties(
            new NumberingLevelReference { Val = ilvl },
            new NumberingId { Val = numId });
        var anchor = style.StyleParagraphProperties.GetFirstChild<SpacingBetweenLines>();
        if (anchor != null)
        {
            style.StyleParagraphProperties.InsertBefore(numberingProperties, anchor);
        }
        else
        {
            style.StyleParagraphProperties.PrependChild(numberingProperties);
        }
    }

    private static Paragraph Paragraph(string text, string? styleId = null)
    {
        var paragraph = new Paragraph();
        var paragraphProperties = new ParagraphProperties();
        if (!string.IsNullOrWhiteSpace(styleId))
        {
            paragraphProperties.Append(new ParagraphStyleId { Val = styleId });
        }

        paragraph.Append(paragraphProperties);
        paragraph.Append(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        return paragraph;
    }

    private static Paragraph ParagraphFromRuns(params string[] texts)
    {
        var paragraph = new Paragraph(new ParagraphProperties());
        foreach (var text in texts)
        {
            paragraph.Append(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        }

        return paragraph;
    }

    private static Paragraph TextBoxParagraph(OpenXmlPartContainer container, string prefix, string text)
    {
        var paragraph = new Paragraph(new ParagraphProperties(), new Run(new Text(prefix) { Space = SpaceProcessingModeValues.Preserve }));
        const string template = """
<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:v="urn:schemas-microsoft-com:vml"
     xmlns:w10="urn:schemas-microsoft-com:office:word">
  <w:pict>
    <v:roundrect id="_x0000_s1027" style="position:absolute;margin-left:10pt;margin-top:10pt;width:120pt;height:30pt" stroked="f">
      <v:textbox inset="0,0,0,0">
        <w:txbxContent>
          <w:p>
            <w:r>
              <w:t>{文本}</w:t>
            </w:r>
          </w:p>
        </w:txbxContent>
      </v:textbox>
      <w10:wrap type="square" anchorx="page" anchory="margin" />
    </v:roundrect>
  </w:pict>
</w:r>
""";
        paragraph.Append(container.CreateUnknownElement(template.Replace("{文本}", text, StringComparison.Ordinal)));
        return paragraph;
    }

    private static Paragraph BoldParagraph(string text)
    {
        return new Paragraph(
            new ParagraphProperties(),
            new Run(
                new RunProperties(new Bold(), new BoldComplexScript()),
                new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static Paragraph TabParagraph(string leftText, string rightText)
    {
        return new Paragraph(
            new ParagraphProperties(),
            new Run(new Text(leftText) { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new TabChar()),
            new Run(new Text(rightText) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static Paragraph NumberedParagraph(string text, int numId, int ilvl)
    {
        var paragraph = Paragraph(text);
        paragraph.ParagraphProperties!.Append(
            new NumberingProperties(
                new NumberingLevelReference { Val = ilvl },
                new NumberingId { Val = numId }));
        return paragraph;
    }

    private static Paragraph OutlineParagraph(string text, int outlineLevel)
    {
        var paragraph = Paragraph(text);
        paragraph.ParagraphProperties!.Append(new OutlineLevel { Val = outlineLevel });
        return paragraph;
    }

    private static TableRow Row(params TableCell[] cells)
    {
        var row = new TableRow();
        foreach (var cell in cells)
        {
            row.Append(cell);
        }
        return row;
    }

    private static TableCell Cell(params string[] texts)
    {
        var paragraph = new Paragraph();
        foreach (var text in texts)
        {
            paragraph.Append(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        }

        return new TableCell(new TableCellProperties(), paragraph);
    }

    private static TableCell ComplexCellWithTabsAndMultipleParagraphs()
    {
        var paragraph1 = new Paragraph(new Run(new Text("1")), new Run(new TabChar()), new Run(new Text("234")));
        var paragraph2 = new Paragraph(new Run(new Text("附加说明")));
        return new TableCell(new TableCellProperties(), paragraph1, paragraph2);
    }

    private static TableCell BookmarkCell(string text, string bookmarkName)
    {
        var paragraph = new Paragraph(
            new BookmarkStart { Id = "1", Name = bookmarkName },
            new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }),
            new BookmarkEnd { Id = "1" });
        return new TableCell(new TableCellProperties(), paragraph);
    }

    private static Table CreateMergedTable()
    {
        return new Table(
            new TableProperties(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }),
            new TableGrid(new GridColumn(), new GridColumn(), new GridColumn()),
            new TableRow(
                new TableCell(
                    new TableCellProperties(new GridSpan { Val = 2 }),
                    new Paragraph(new Run(new Text("横向合并")))),
                new TableCell(new Paragraph(new Run(new Text("右侧单元格"))))),
            new TableRow(
                new TableCell(
                    new TableCellProperties(new VerticalMerge { Val = MergedCellValues.Restart }),
                    new Paragraph(new Run(new Text("纵向起点")))),
                new TableCell(new Paragraph(new Run(new Text("普通格")))),
                new TableCell(new Paragraph(new Run(new Text("普通格二"))))),
            new TableRow(
                new TableCell(
                    new TableCellProperties(new VerticalMerge()),
                    new Paragraph(new Run(new Text("纵向延续")))),
                new TableCell(new Paragraph(new Run(new Text("普通格三")))),
                new TableCell(new Paragraph(new Run(new Text("普通格四"))))));
    }

    private static IEnumerable<Paragraph> CreateComplexSignoffParagraphs()
    {
        yield return new Paragraph(
            new ParagraphProperties(),
            new BookmarkStart { Id = "2", Name = "落款书签" },
            new Run(new Text("四川华信(集团)会计师事务所")),
            new BookmarkEnd { Id = "2" });
        yield return new Paragraph(new ParagraphProperties(), new Run(new Text("（特殊普通合伙）")));
        yield return new Paragraph(
            new ParagraphProperties(),
            new Hyperlink(new Run(new Text("中国·成都"))) { Anchor = "落款书签" });
        yield return new Paragraph(new ParagraphProperties(), new Run(new Text("中国注册会计师：张三")));
        yield return new Paragraph(new ParagraphProperties(), new Run(new Text("中国注册会计师：李四")));
        yield return new Paragraph(new ParagraphProperties(), new Run(new Text("二〇二六年十月二日")));
    }

    private static Table BuildCoverKeywordTable(bool splitRuns)
    {
        var table = new Table(
            new TableProperties(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }),
            new TableGrid(new GridColumn(), new GridColumn()));

        TableCell BuildCell(params string[] parts)
        {
            var paragraph = new Paragraph();
            foreach (var part in parts)
            {
                paragraph.Append(new Run(new Text(part) { Space = SpaceProcessingModeValues.Preserve }));
            }
            return new TableCell(new TableCellProperties(), paragraph);
        }

        if (splitRuns)
        {
            table.Append(
                Row(BuildCell("会计师", "事务所"), BuildCell("地", "址")),
                Row(BuildCell("电", "\t", "话"), BuildCell("传", "真")));
        }
        else
        {
            table.Append(
                Row(BuildCell("会计师事务所"), BuildCell("地址")),
                Row(BuildCell("电话"), BuildCell("传真")));
        }

        return table;
    }

    private static SectionProperties BuildSectionProperties(string? headerId = null, string? footerId = null)
    {
        var sectionProperties = new SectionProperties();

        if (!string.IsNullOrWhiteSpace(headerId))
        {
            sectionProperties.Append(new HeaderReference { Id = headerId, Type = HeaderFooterValues.Default });
        }

        if (!string.IsNullOrWhiteSpace(footerId))
        {
            sectionProperties.Append(new FooterReference { Id = footerId, Type = HeaderFooterValues.Default });
        }

        sectionProperties.Append(
            new PageSize { Width = 11906U, Height = 16838U },
            new PageMargin
            {
                Top = 1440,
                Bottom = 1440,
                Left = (UInt32Value)1440U,
                Right = (UInt32Value)1440U,
                Header = (UInt32Value)720U,
                Footer = (UInt32Value)720U,
                Gutter = (UInt32Value)0U
            });

        return sectionProperties;
    }

    private static SectionProperties BuildBodySectionProperties(string? headerId = null, string? footerId = null)
    {
        var sectionProperties = BuildSectionProperties(headerId, footerId);
        var pageMargin = sectionProperties.GetFirstChild<PageMargin>();
        if (pageMargin != null)
        {
            sectionProperties.InsertAfter(new PageNumberType { Start = 1 }, pageMargin);
        }
        else
        {
            sectionProperties.Append(new PageNumberType { Start = 1 });
        }
        return sectionProperties;
    }

    private static SectionProperties BuildLandscapeSectionProperties(string? headerId = null, string? footerId = null)
    {
        var sectionProperties = BuildBodySectionProperties(headerId, footerId);
        var pageSize = sectionProperties.GetFirstChild<PageSize>();
        if (pageSize != null)
        {
            pageSize.Width = 16838U;
            pageSize.Height = 11906U;
            pageSize.Orient = PageOrientationValues.Landscape;
        }

        return sectionProperties;
    }

    private static void AppendBodySectionProperties(Body body)
    {
        body.Append(BuildBodySectionProperties());
    }

    private static void AppendPaddingBody(Body body, string prefix)
    {
        var chunk = "本段用于让最小矩阵样本稳定进入正式排版主链，同时覆盖正文可见文本统计、标题识别、编号保留、表格清洗和验证链路。";
        for (var i = 1; i <= 3; i++)
        {
            body.Append(Paragraph($"{prefix}{i}：{chunk}{chunk}"));
        }
    }

    private static string BuildExpectedMarkdown()
    {
        var builder = new StringBuilder();
        builder.AppendLine("# 最小 DOCX 结构矩阵说明");
        builder.AppendLine();
        builder.AppendLine("- `标题矩阵`：覆盖文本前缀、样式链编号、段落直接编号、仅大纲标题");
        builder.AppendLine("- `正文编号矩阵`：覆盖正文式四级、普通说明列表、样式链编号正文");
        builder.AppendLine("- `表格复杂文本矩阵`：覆盖首列正整数字符串、拆分数字、复杂结构单元格、百分比");
        builder.AppendLine("- `封面矩阵`：覆盖封面关键词拆分运行块和分节");
        builder.AppendLine("- `纯封面矩阵`：覆盖低于 200 字的纯封面文档");
        builder.AppendLine("- `落款矩阵`：覆盖事务所名、城市、注册会计师、日期模板占位");
        builder.AppendLine("- `共享编号模板防误伤矩阵`：覆盖短标题与长编号正文共用 numbering 模板");
        builder.AppendLine("- `页眉页脚分节矩阵`：覆盖封面节与正文节共享页眉页脚部件");
        builder.AppendLine("- `标题坑点矩阵`：覆盖标题 run 拆分、样式链挂编号但可见文本不带前缀");
        builder.AppendLine("- `表格坑点矩阵`：覆盖多段复杂单元格、书签等附着结构、拆分数字 run");
        builder.AppendLine("- `页眉页脚坑点矩阵`：覆盖横竖页混排、旧页眉污染、带域代码页脚");
        builder.AppendLine("- `落款坑点矩阵`：覆盖旧空格、旧 tab、旧粗体污染");
        builder.AppendLine("- `非目标区域保护矩阵`：覆盖前导零文本、表格内落款关键字、共享编号长正文保护");
        builder.AppendLine("- `文本框坑点矩阵`：覆盖文本框文字不在普通正文树里的情况");
        builder.AppendLine("- `脚注尾注坑点矩阵`：覆盖脚注部件、尾注部件与正文引用");
        builder.AppendLine("- `编号重启坑点矩阵`：覆盖编号实例重启结构");
        builder.AppendLine("- `合并单元格坑点矩阵`：覆盖横向合并与纵向合并");
        builder.AppendLine("- `落款复杂结构坑点矩阵`：覆盖落款区里的书签和链接混排");
        return builder.ToString();
    }
}
