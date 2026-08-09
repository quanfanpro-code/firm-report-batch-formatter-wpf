using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FirmFormatter.OpenXml.Contracts;
using FirmFormatter.OpenXml.Core;
using Xunit;

namespace 文档快照导出测试;

public sealed class 复杂结构快照测试
{
    [Fact]
    public void 导出快照时_应包含复杂结构统计()
    {
        using var stream = new MemoryStream();
        using var word = CreateComplexDocument(stream);
        var request = new RequestContract
        {
            InputPath = "输入.docx",
            OutputPath = "输出.docx",
            ScenarioName = "复杂结构",
            RunSource = "测试"
        };

        var context = new FirmDocumentClassifier().BuildContext(word, request);
        var json = new DocumentSnapshotService().ExportAsJson(word, "输出.docx", context);
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;

        Assert.True(root.GetProperty("正文可见字数").GetInt32() > 0);
        Assert.True(root.GetProperty("表格详情").GetArrayLength() > 0);

        var complex = root.GetProperty("复杂结构");
        Assert.Equal(1, complex.GetProperty("文本框数").GetInt32());
        Assert.Equal(1, complex.GetProperty("脚注数").GetInt32());
        Assert.Equal(1, complex.GetProperty("尾注数").GetInt32());
        Assert.Equal(1, complex.GetProperty("编号重启数").GetInt32());
        Assert.Equal(1, complex.GetProperty("横向合并数").GetInt32());
        Assert.Equal(2, complex.GetProperty("纵向合并数").GetInt32());
        Assert.Equal(2, complex.GetProperty("书签数").GetInt32());
        Assert.Equal(2, complex.GetProperty("超链接数").GetInt32());
        Assert.Equal(1, complex.GetProperty("域代码数").GetInt32());
        Assert.Equal(2, complex.GetProperty("落款复杂结构数").GetInt32());
        Assert.True(complex.GetProperty("文本框摘要").GetString()!.Contains("文本框里的话", StringComparison.Ordinal));
        Assert.True(complex.GetProperty("脚注尾注摘要").GetString()!.Contains("脚注内容", StringComparison.Ordinal));
    }

    private static WordprocessingDocument CreateComplexDocument(MemoryStream stream)
    {
        var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var mainPart = word.AddMainDocumentPart();
        var body = new Body();
        mainPart.Document = new Document(body);

        body.Append(CreateTextBoxParagraph(mainPart));
        body.Append(CreateBookmarkAndHyperlinkParagraph());
        body.Append(CreateFieldCodeParagraph());
        body.Append(NumberedParagraph("第一段编号正文", 92, 0));
        body.Append(NumberedParagraph("重启后的编号正文", 93, 0));
        body.Append(CreateMergedTable());
        body.Append(new Paragraph(new Run(new Text(new string('正', 260)))));
        body.Append(CreateComplexSignoffParagraphs());
        body.Append(new SectionProperties());

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

        var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new Numbering(
            new AbstractNum(
                new Level(
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%1." },
                    new LevelJustification { Val = LevelJustificationValues.Left })
                { LevelIndex = 0 })
            { AbstractNumberId = 91, MultiLevelType = new MultiLevelType { Val = MultiLevelValues.HybridMultilevel } },
            new NumberingInstance(new AbstractNumId { Val = 91 }) { NumberID = 92 },
            new NumberingInstance(
                new AbstractNumId { Val = 91 },
                new LevelOverride(
                    new StartOverrideNumberingValue { Val = 1 })
                { LevelIndex = 0 })
            { NumberID = 93 });
        numberingPart.Numbering.Save();

        mainPart.Document.Save();
        return word;
    }

    private static Paragraph CreateTextBoxParagraph(OpenXmlPartContainer container)
    {
        const string textBoxXml = """
<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:v="urn:schemas-microsoft-com:vml"
     xmlns:w10="urn:schemas-microsoft-com:office:word">
  <w:pict>
    <v:roundrect id="_x0000_s1027" style="position:absolute;margin-left:10pt;margin-top:10pt;width:120pt;height:30pt" stroked="f">
      <v:textbox inset="0,0,0,0">
        <w:txbxContent>
          <w:p>
            <w:r>
              <w:t>文本框里的话</w:t>
            </w:r>
          </w:p>
        </w:txbxContent>
      </v:textbox>
      <w10:wrap type="square" anchorx="page" anchory="margin" />
    </v:roundrect>
  </w:pict>
</w:r>
""";

        var paragraph = new Paragraph(new Run(new Text("正文前缀")));
        paragraph.Append(container.CreateUnknownElement(textBoxXml));
        return paragraph;
    }

    private static Paragraph CreateBookmarkAndHyperlinkParagraph()
    {
        return new Paragraph(
            new BookmarkStart { Id = "1", Name = "书签一" },
            new Hyperlink(new Run(new Text("链接文字"))) { Anchor = "书签一" },
            new BookmarkEnd { Id = "1" });
    }

    private static Paragraph CreateFieldCodeParagraph()
    {
        return new Paragraph(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" PAGE ")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("1")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
    }

    private static Paragraph NumberedParagraph(string text, int numId, int ilvl)
    {
        return new Paragraph(
            new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = ilvl },
                    new NumberingId { Val = numId })),
            new Run(new Text(text)));
    }

    private static IEnumerable<Paragraph> CreateComplexSignoffParagraphs()
    {
        yield return new Paragraph(
            new BookmarkStart { Id = "2", Name = "落款书签" },
            new Run(new Text("四川华信(集团)会计师事务所")),
            new BookmarkEnd { Id = "2" });
        yield return new Paragraph(new Run(new Text("（特殊普通合伙）")));
        yield return new Paragraph(new Hyperlink(new Run(new Text("中国·成都"))) { Anchor = "落款书签" });
        yield return new Paragraph(new Run(new Text("中国注册会计师：张三")));
        yield return new Paragraph(new Run(new Text("中国注册会计师：李四")));
        yield return new Paragraph(new Run(new Text("二〇二六年十月二日")));
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
}
