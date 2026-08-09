using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FirmFormatter.OpenXml.Core;

public sealed class TableService
{
    private const uint OuterBorderSize = 6U;
    private const uint InnerBorderSize = 6U;

    // 以下为表格专用排版参数，取值应与 FirmRuleProfile（正文字号HalfPoint / 中文字体 / 西文字体）保持一致；
    // FirmRuleProfile 暂未收纳字符间距等表格参数，如需调整请两边同步修改
    private const string 表格字号HalfPoint = "24";
    private const int 表格字符间距紧缩Twips = -20; // twips 单位，-20 表示紧缩 1 磅
    private const string 表格中文字体 = "宋体";
    private const string 表格西文字体 = "Times New Roman";

    public void Apply(WordprocessingDocument word, bool hasCover)
    {
        var body = word.MainDocumentPart?.Document?.Body;
        if (body is null) return;

        var sectionIndex = 0;
        foreach (var child in body.Elements())
        {
            if (child is Paragraph p && p.ParagraphProperties?.GetFirstChild<SectionProperties>() is not null)
            {
                sectionIndex++;
            }

            if (child is not Table table) continue;
            if (hasCover && sectionIndex == 0) continue;
            // 封面表判定只限第一节：正文/落款附近含"会计师事务所、地址、电话、传真"的信息表不应被误判为封面表而跳过格式化
            if (sectionIndex == 0 && IsCoverTable(table)) continue;
            ApplyTable(table);
        }
    }

    private static bool IsCoverTable(Table table)
    {
        var t = OpenXmlHelper.NormalizeText(OpenXmlHelper.提取可见文本(table));
        return t.Contains("会计师事务所") && t.Contains("地址") && t.Contains("电话") && t.Contains("传真");
    }

    private static void ApplyTable(Table table)
    {
        var tableProps = table.GetFirstChild<TableProperties>();
        if (tableProps is null)
        {
            tableProps = new TableProperties();
            table.PrependChild(tableProps);
        }

        tableProps.GetFirstChild<TableStyle>()?.Remove();
        tableProps.GetFirstChild<TableLook>()?.Remove();
        tableProps.GetFirstChild<TableStyleRowBandSize>()?.Remove();
        tableProps.GetFirstChild<TableStyleColumnBandSize>()?.Remove();
        tableProps.GetFirstChild<TableCellSpacing>()?.Remove();
        tableProps.TableWidth = new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" };
        tableProps.TableBorders = new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = OuterBorderSize, Color = "000000" },
            new LeftBorder { Val = BorderValues.Nil },
            new BottomBorder { Val = BorderValues.Single, Size = OuterBorderSize, Color = "000000" },
            new RightBorder { Val = BorderValues.Nil },
            new InsideHorizontalBorder { Val = BorderValues.Dotted, Size = InnerBorderSize, Color = "000000" },
            new InsideVerticalBorder { Val = BorderValues.Dotted, Size = InnerBorderSize, Color = "000000" }
        );

        var rows = table.Elements<TableRow>().ToList();
        var sequenceColumns = rows.Count == 0
            ? []
            : rows[0].Elements<TableCell>()
                .Select((cell, index) => new
                {
                    Index = index,
                    Header = OpenXmlHelper.NormalizeText(OpenXmlHelper.提取可见文本(cell))
                })
                .Where(item => string.Equals(item.Header, "序号", StringComparison.Ordinal))
                .Select(item => item.Index)
                .ToHashSet();
        if (rows.Count > 0)
        {
            var firstRow = rows[0];
            var trPr = firstRow.GetFirstChild<TableRowProperties>();
            if (trPr == null)
            {
                trPr = new TableRowProperties();
                firstRow.PrependChild(trPr);
            }
            if (trPr.GetFirstChild<TableHeader>() == null)
            {
                InsertTableHeaderInSchemaOrder(trPr);
            }
        }
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var row = rows[rowIndex];
            row.GetFirstChild<TableRowProperties>()?.GetFirstChild<TableCellSpacing>()?.Remove();
            var cells = row.Elements<TableCell>().ToList();
            for (int colIndex = 0; colIndex < cells.Count; colIndex++)
            {
                var cell = cells[colIndex];
                var tcPr = cell.GetFirstChild<TableCellProperties>();
                if (tcPr == null)
                {
                    tcPr = new TableCellProperties();
                    cell.PrependChild(tcPr);
                }

                var tcBorders = tcPr.GetFirstChild<TableCellBorders>() ?? new TableCellBorders();
                tcBorders.TopBorder = BuildTopBorder(rowIndex == 0);
                tcBorders.BottomBorder = BuildBottomBorder(rowIndex == rows.Count - 1);
                tcBorders.LeftBorder = BuildLeftBorder(colIndex == 0);
                tcBorders.RightBorder = BuildRightBorder(colIndex == cells.Count - 1);
                tcPr.TableCellBorders = tcBorders;

                tcPr.TableCellMargin = new TableCellMargin
                {
                    TopMargin = new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                    BottomMargin = new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa }
                };
                tcPr.TableCellVerticalAlignment = new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center };

                var text = OpenXmlHelper.提取可见文本(cell).Trim();
                var formatted = FormatCellDisplayText(rowIndex, sequenceColumns.Contains(colIndex), text, out var isPureNumeric);
                if (formatted is not null && 单元格可安全重写(cell))
                {
                    ReplaceCellTextPreservingStructure(cell, formatted);
                }

                foreach (var p in cell.Elements<Paragraph>())
                {
                    OpenXmlHelper.EnsureParagraphProperties(p);
                    // 彻底清除所有可能的缩进属性，防止继承原有的奇葩缩进
                    p.ParagraphProperties!.Indentation = new Indentation
                    {
                        Left = "0",
                        Right = "0",
                        FirstLine = "0",
                        Start = "0",
                        LeftChars = 0,
                        RightChars = 0,
                        FirstLineChars = 0,
                        StartCharacters = 0,
                        Hanging = null,
                        HangingChars = null
                    };

                    // 行距：单倍行距，彻底覆盖原有奇葩行距
                    p.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines
                    {
                        Before = "0",
                        After = "0",
                        BeforeLines = 0,
                        AfterLines = 0,
                        BeforeAutoSpacing = false,
                        AfterAutoSpacing = false,
                        LineRule = LineSpacingRuleValues.Auto,
                        Line = "240"
                    };

                    p.ParagraphProperties.GetFirstChild<Tabs>()?.Remove();
                    p.ParagraphProperties.GetFirstChild<NumberingProperties>()?.Remove();
                    if (p.ParagraphProperties.OutlineLevel != null) p.ParagraphProperties.OutlineLevel.Remove();
                    if (p.ParagraphProperties.KeepNext != null) p.ParagraphProperties.KeepNext.Remove();

                    if (rowIndex == 0)
                    {
                        p.ParagraphProperties.Justification = new Justification { Val = JustificationValues.Center };
                    }
                    else if (colIndex == 0)
                    {
                        p.ParagraphProperties.Justification = new Justification { Val = JustificationValues.Left };
                    }
                    else if (isPureNumeric)
                    {
                        p.ParagraphProperties.Justification = new Justification { Val = JustificationValues.Right };
                    }
                    else
                    {
                        p.ParagraphProperties.Justification = new Justification { Val = JustificationValues.Center };
                    }

                    foreach (var run in p.Descendants<Run>())
                    {
                        run.RunProperties ??= new RunProperties();
                        run.RunProperties.FontSize = new FontSize { Val = 表格字号HalfPoint };

                        // 移除任何字体缩放，恢复默认100%
                        var scale = run.RunProperties.GetFirstChild<CharacterScale>();
                        if (scale != null) scale.Remove();

                        var fitText = run.RunProperties.GetFirstChild<FitText>();
                        if (fitText != null) fitText.Remove();

                        var position = run.RunProperties.GetFirstChild<Position>();
                        if (position != null) position.Remove();

                        var runStyle = run.RunProperties.GetFirstChild<RunStyle>();
                        if (runStyle != null) runStyle.Remove();

                        // 设置字符间距紧缩，Val为负值（twips单位，-20 表示紧缩 1 磅）
                        run.RunProperties.Spacing = new Spacing { Val = 表格字符间距紧缩Twips };

                        // 清除加粗、斜体、下划线
                        run.RunProperties.Bold = new Bold { Val = false };
                        run.RunProperties.BoldComplexScript = new BoldComplexScript { Val = false };
                        run.RunProperties.Italic = new Italic { Val = false };
                        run.RunProperties.ItalicComplexScript = new ItalicComplexScript { Val = false };

                        var underline = run.RunProperties.GetFirstChild<Underline>();
                        if (underline != null) underline.Remove();

                        OpenXmlHelper.SetRunFonts(run.RunProperties, 表格中文字体, 表格西文字体);
                    }
                }
            }
        }
    }

    private static string? FormatCellDisplayText(int rowIndex, bool isSequenceColumn, string rawText, out bool isPureNumeric)
    {
        isPureNumeric = false;

        // 首行恒为表头行（ApplyTable 已为 rows[0] 加 TableHeader 并居中），
        // 表头不做数字格式化，避免"2023"这类纯数字表头被改成"2,023.00"
        if (rowIndex == 0) return null;

        // 表头明确为“序号”的整列保持原始显示，包括 0、前导零和带点编号。
        if (isSequenceColumn) return null;

        var formatted = FormatPureNumericCell(rawText);
        isPureNumeric = formatted != null;
        return formatted;
    }

    private static string? FormatPureNumericCell(string rawText)
    {
        var raw = rawText.Replace("\u3000", "").Replace(" ", "").Replace("（", "(").Replace("）", ")").Replace("[", "(").Replace("]", ")").Replace("％", "%").Replace("，", ","); // 全角 ％→%、，→,（全角逗号随后作为千分位剔除）
        if (raw.Length == 0) return null;
        if (raw == "-") return "";

        var bracket = raw.StartsWith("(") && raw.EndsWith(")");
        var normalized = bracket ? raw[1..^1] : raw;
        var isPercent = normalized.EndsWith("%");
        var numericPart = isPercent ? normalized[..^1] : normalized;
        numericPart = numericPart.Replace(",", "");

        if (!Regex.IsMatch(numericPart, @"^[+-]?\d+(\.\d+)?$")) return null;
        if (!decimal.TryParse(numericPart, NumberStyles.Any, CultureInfo.InvariantCulture, out var value)) return null;

        if (bracket) value = -Math.Abs(value);
        if (value == 0) return "";

        var abs = Math.Abs(value);
        var formatted = abs.ToString("N2", CultureInfo.InvariantCulture);
        if (isPercent) formatted += "%";

        if (value < 0)
        {
            return $"-{formatted}";
        }
        return formatted;
    }

    private static void ReplaceCellTextPreservingStructure(TableCell cell, string formatted)
    {
        var texts = cell.Descendants<Text>().ToList();
        if (texts.Count == 0)
        {
            cell.RemoveAllChildren<Paragraph>();
            cell.AppendChild(new Paragraph(new Run(new Text(formatted))));
            return;
        }

        texts[0].Text = formatted;
        texts[0].Space = SpaceProcessingModeValues.Preserve;

        for (var i = 1; i < texts.Count; i++)
        {
            texts[i].Remove();
        }
    }

    private static bool 单元格可安全重写(TableCell cell)
    {
        if (cell.Elements<Paragraph>().Count() != 1) return false;
        if (cell.Descendants<SimpleField>().Any()) return false;
        if (cell.Descendants<FieldCode>().Any()) return false;
        if (cell.Descendants<Break>().Any()) return false;
        if (cell.Descendants<TabChar>().Any()) return false;
        if (cell.Descendants<Drawing>().Any()) return false;
        if (cell.Descendants<Hyperlink>().Any()) return false;
        if (cell.Descendants<InsertedRun>().Any()) return false;
        if (cell.Descendants<DeletedRun>().Any()) return false;
        if (cell.Descendants<OfficeMath>().Any()) return false;
        // Descendants<Table> 只会返回单元格内的嵌套表（cell.Parent 是 TableRow，原 x != cell.Parent 比较恒真），直接判存在即可
        if (cell.Descendants<Table>().Any()) return false;

        return true;
    }

    private static TopBorder BuildTopBorder(bool isFirstRow)
    {
        return new TopBorder
        {
            Val = isFirstRow ? BorderValues.Single : BorderValues.Nil,
            Size = isFirstRow ? OuterBorderSize : InnerBorderSize,
            Color = "000000"
        };
    }

    private static BottomBorder BuildBottomBorder(bool isLastRow)
    {
        return new BottomBorder
        {
            Val = isLastRow ? BorderValues.Single : BorderValues.Dotted,
            Size = isLastRow ? OuterBorderSize : InnerBorderSize,
            Color = "000000"
        };
    }

    private static LeftBorder BuildLeftBorder(bool isFirstColumn)
    {
        return new LeftBorder
        {
            Val = isFirstColumn ? BorderValues.Nil : BorderValues.Dotted,
            Size = InnerBorderSize,
            Color = "000000"
        };
    }

    private static RightBorder BuildRightBorder(bool isLastColumn)
    {
        return new RightBorder
        {
            Val = isLastColumn ? BorderValues.Nil : BorderValues.Nil,
            Size = InnerBorderSize,
            Color = "000000"
        };
    }

    private static void InsertTableHeaderInSchemaOrder(TableRowProperties trPr)
    {
        var anchor = trPr.Elements<OpenXmlElement>()
            .FirstOrDefault(element =>
                element is TableCellSpacing ||
                element is Justification);

        if (anchor != null)
        {
            trPr.InsertBefore(new TableHeader(), anchor);
            return;
        }

        trPr.AppendChild(new TableHeader());
    }
}
