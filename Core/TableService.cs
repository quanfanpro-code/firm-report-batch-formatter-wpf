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
    private static readonly string[] 数字格式豁免列表头关键词 =
    [
        "序号", "编号", "代码", "号码",
        "账号", "识别号",
        "年份", "年度", "月份", "季度", "日期", "账龄", "期数", "页码",
        "数量", "人数", "户数", "件数", "台数"
    ];

    private readonly FirmRuleProfile _ruleProfile = FirmRuleProfile.Default;
    // 字符间距仍是表格专用参数，其他通用字体字号直接使用统一规则来源。
    private const int 表格字符间距紧缩Twips = -20; // twips 单位，-20 表示紧缩 1 磅
    public int 已格式化数字单元格数 { get; private set; }
    public int 发生两位小数舍入的单元格数 { get; private set; }

    public void Apply(WordprocessingDocument word, bool hasCover, CancellationToken cancellationToken = default)
    {
        已格式化数字单元格数 = 0;
        发生两位小数舍入的单元格数 = 0;
        var body = word.MainDocumentPart?.Document?.Body;
        if (body is null) return;

        var hasDedicatedCoverSection = hasCover && OpenXmlHelper.收集分节(body).Count > 1;
        var sectionIndex = 0;
        foreach (var child in body.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (child is Paragraph p && p.ParagraphProperties?.GetFirstChild<SectionProperties>() is not null)
            {
                sectionIndex++;
            }

            if (child is not Table table) continue;
            if (hasDedicatedCoverSection && sectionIndex == 0) continue;
            // 封面表判定只限第一节：正文/落款附近含"会计师事务所、地址、电话、传真"的信息表不应被误判为封面表而跳过格式化
            if (sectionIndex == 0 && IsCoverTable(table)) continue;
            ApplyTable(table, cancellationToken);
        }
    }

    private static bool IsCoverTable(Table table)
    {
        var t = OpenXmlHelper.NormalizeText(OpenXmlHelper.提取可见文本(table));
        return t.Contains("会计师事务所") && t.Contains("地址") && t.Contains("电话") && t.Contains("传真");
    }

    private void ApplyTable(Table table, CancellationToken cancellationToken)
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
        var headerRowCount = GetHeaderRowCount(rows);
        var numberFormatExemptColumns = GetNumberFormatExemptColumns(rows, headerRowCount);
        for (var headerRowIndex = 0; headerRowIndex < headerRowCount; headerRowIndex++)
        {
            var headerRow = rows[headerRowIndex];
            var trPr = headerRow.GetFirstChild<TableRowProperties>();
            if (trPr == null)
            {
                trPr = new TableRowProperties();
                headerRow.PrependChild(trPr);
            }
            if (trPr.GetFirstChild<TableHeader>() == null)
            {
                InsertTableHeaderInSchemaOrder(trPr);
            }
        }
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = rows[rowIndex];
            row.GetFirstChild<TableRowProperties>()?.GetFirstChild<TableCellSpacing>()?.Remove();
            var cells = row.Elements<TableCell>().ToList();
            var logicalColIndex = row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;
            for (int colIndex = 0; colIndex < cells.Count; colIndex++)
            {
                var cell = cells[colIndex];
                var currentLogicalColIndex = logicalColIndex;
                var logicalSpan = Math.Max(1, cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
                logicalColIndex += logicalSpan;
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
                var isNumberFormatExemptColumn = Enumerable.Range(currentLogicalColIndex, logicalSpan)
                    .Any(numberFormatExemptColumns.Contains);
                var formatted = FormatCellDisplayText(rowIndex, headerRowCount, isNumberFormatExemptColumn, text, out var isPureNumeric, out var wasRounded);
                if (formatted is not null && 单元格可安全重写(cell))
                {
                    ReplaceCellTextPreservingStructure(cell, formatted);
                    已格式化数字单元格数++;
                    if (wasRounded) 发生两位小数舍入的单元格数++;
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
                    if (p.ParagraphProperties.OutlineLevel != null) p.ParagraphProperties.OutlineLevel.Remove();
                    if (p.ParagraphProperties.KeepNext != null) p.ParagraphProperties.KeepNext.Remove();

                    if (rowIndex < headerRowCount)
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
                        if (run.Descendants<FootnoteReference>().Any() || run.Descendants<EndnoteReference>().Any())
                        {
                            continue;
                        }

                        run.RunProperties ??= new RunProperties();
                        run.RunProperties.FontSize = new FontSize { Val = _ruleProfile.正文字号HalfPoint };

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

                        OpenXmlHelper.SetRunFonts(run.RunProperties, _ruleProfile.中文字体, _ruleProfile.西文字体);
                    }
                }
            }
        }
    }

    private static string? FormatCellDisplayText(int rowIndex, int headerRowCount, bool isNumberFormatExemptColumn, string rawText, out bool isPureNumeric, out bool wasRounded)
    {
        isPureNumeric = false;
        wasRounded = false;

        // 首行恒为表头行（ApplyTable 已为 rows[0] 加 TableHeader 并居中），
        // 表头不做数字格式化，避免"2023"这类纯数字表头被改成"2,023.00"
        if (rowIndex < headerRowCount) return null;

        // 标识、期间和数量类列保持原始显示，包括 0、前导零和带点编号。
        if (isNumberFormatExemptColumn) return null;

        var formatted = FormatPureNumericCell(rawText, out wasRounded);
        isPureNumeric = formatted != null;
        return formatted;
    }

    internal static int GetHeaderRowCount(IReadOnlyList<TableRow> rows)
    {
        if (rows.Count == 0) return 0;

        var count = 1;
        while (count < rows.Count && rows[count].GetFirstChild<TableRowProperties>()?.GetFirstChild<TableHeader>() is not null)
        {
            count++;
        }

        if (count == 1 && rows.Count > 1 && rows[0].Elements<TableCell>()
                .Select(cell => cell.TableCellProperties)
                .Any(properties => properties is not null
                    && ((properties.GridSpan?.Val?.Value ?? 1) > 1 || properties.VerticalMerge is not null)))
        {
            count = 2;
        }

        return count;
    }

    private static IEnumerable<(int Index, string Header)> EnumerateLogicalCells(TableRow row)
    {
        var logicalIndex = row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;
        foreach (var cell in row.Elements<TableCell>())
        {
            var header = OpenXmlHelper.NormalizeText(OpenXmlHelper.提取可见文本(cell));
            var span = Math.Max(1, cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
            for (var offset = 0; offset < span; offset++)
            {
                yield return (logicalIndex + offset, header);
            }
            logicalIndex += span;
        }
    }

    private static HashSet<int> GetNumberFormatExemptColumns(IReadOnlyList<TableRow> rows, int headerRowCount)
    {
        var leafHeaders = new Dictionary<int, string>();
        foreach (var row in rows.Take(headerRowCount))
        {
            foreach (var item in EnumerateLogicalCells(row))
            {
                if (!string.IsNullOrWhiteSpace(item.Header))
                {
                    leafHeaders[item.Index] = item.Header;
                }
            }
        }

        return leafHeaders
            .Where(item => IsNumberFormatExemptHeader(item.Value))
            .Select(item => item.Key)
            .ToHashSet();
    }

    private static bool IsNumberFormatExemptHeader(string header)
    {
        return 数字格式豁免列表头关键词.Any(keyword => header.Contains(keyword, StringComparison.Ordinal))
            || header.EndsWith("账户", StringComparison.Ordinal)
            || header.Contains("账户号", StringComparison.Ordinal);
    }

    private static string? FormatPureNumericCell(string rawText, out bool wasRounded)
    {
        wasRounded = false;
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
        wasRounded = decimal.Round(value, 2, MidpointRounding.ToEven) != value;

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
        if (cell.Descendants<FootnoteReference>().Any() || cell.Descendants<EndnoteReference>().Any()) return false;
        if (cell.Descendants<Vanish>().Any() || cell.Descendants<WebHidden>().Any()) return false;
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
            .FirstOrDefault(element => element.LocalName is
                "wBefore" or "wAfter" or "tblCellSpacing" or "jc" or
                "ins" or "del" or "conflictIns" or "conflictDel" or "trPrChange");

        if (anchor != null)
        {
            trPr.InsertBefore(new TableHeader(), anchor);
            return;
        }

        trPr.AppendChild(new TableHeader());
    }
}
