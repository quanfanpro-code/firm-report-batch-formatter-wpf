using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FirmFormatter.OpenXml.Core;

public static class OpenXmlHelper
{
    public static List<SectionProperties> 收集分节(Body? body)
    {
        if (body == null) return [];

        var sections = new List<SectionProperties>();
        var seen = new HashSet<SectionProperties>();
        foreach (var paragraph in body.Elements<Paragraph>())
        {
            var section = paragraph.ParagraphProperties?.GetFirstChild<SectionProperties>();
            if (section != null && seen.Add(section)) sections.Add(section);
        }

        var bodySection = body.GetFirstChild<SectionProperties>();
        if (bodySection != null && seen.Add(bodySection)) sections.Add(bodySection);
        return sections;
    }

    public static readonly Regex DateRegex = new(@"(?:20\d{2}|[〇零一二三四五六七八九十]{2,4})年(?:[〇零一二三四五六七八九十\d]{1,2}|十[一二三四五六七八九十]|[一二三四五六七八九十])月(?:[〇零一二三四五六七八九十\d]{1,2}|十[一二三四五六七八九十]|[一二三四五六七八九十])日", RegexOptions.Compiled);
    // 以下三个标题正则与 ParagraphService 中的定义同源重复，改为 internal 供 ParagraphService 复用
    internal static readonly Regex 一级标题文本 = new(@"^[一二三四五六七八九十]+、", RegexOptions.Compiled);
    internal static readonly Regex 二级标题文本 = new(@"^[（(][一二三四五六七八九十]+[)）]", RegexOptions.Compiled);
    internal static readonly Regex 三级标题文本 = new(@"^\d{1,3}[、\.．]", RegexOptions.Compiled);
    private static readonly Regex 正文式四级文本 = new(@"^[（(]\d+[)）](?:[、\.．])?", RegexOptions.Compiled);
    private static readonly Regex 二级标题编号模板 = new(@"^[（(]%\d+[)）](、)?$", RegexOptions.Compiled);
    private static readonly Regex 三级标题编号模板 = new(@"^%\d+[、\.．]$", RegexOptions.Compiled);
    private static readonly Regex 文本框片段 = new(@"<(?:(?:\w+):)?txbxContent\b[\s\S]*?</(?:(?:\w+):)?txbxContent>", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex 文本框文字节点 = new(@"<(?:(?:\w+):)?t\b[^>]*>(.*?)</(?:(?:\w+):)?t>", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public static string NormalizeText(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return string.Empty;
        var span = text.AsSpan();
        // 病理级超长段落用 stackalloc 会栈溢出直接崩进程，超过阈值改堆分配
        Span<char> buffer = span.Length > 8192 ? new char[span.Length] : stackalloc char[span.Length];
        var written = 0;
        foreach (var c in span)
        {
            if (c == '\u3000' || c == ' ' || c == '\t') continue;
            buffer[written++] = c;
        }
        var result = buffer[..written].Trim();
        return result.Length == 0 ? string.Empty : result.ToString();
    }

    public static string ParagraphText(Paragraph p)
    {
        var text = new StringBuilder();
        追加可见文本(p, text, p);
        return text.ToString();
    }

    public static int 统计正文可见有效文本字数(Body? body)
    {
        if (body == null) return 0;
        return 提取可见文本(body).Count(ch => !char.IsWhiteSpace(ch) && ch != '\u3000');
    }

    public static bool 正文包含有效可见内容(Body? body)
    {
        return 统计正文可见有效文本字数(body) > 0;
    }

    public static int ResolveHeadingLevelByVisibleNumbering(MainDocumentPart? mainPart, Paragraph p)
    {
        var paragraphText = NormalizeText(ParagraphText(p));
        if (paragraphText.Length > 30) return 0;

        var numPr = ResolveParagraphNumberingProperties(mainPart, p);
        var numId = numPr?.GetFirstChild<NumberingId>()?.Val?.Value;
        var level = ResolveNumberingLevel(mainPart, p);
        if (level == null) return 0;

        var levelText = level.LevelText?.Val?.Value ?? string.Empty;
        if (string.IsNullOrWhiteSpace(levelText)) return 0;

        var normalized = NormalizeText(levelText);
        var levelIndex = level.LevelIndex?.Value;
        if (是括号汉字二级标题编号(level, normalized) && levelIndex == 1)
        {
            return 2;
        }

        if (normalized.StartsWith("%1、") && levelIndex == 0)
        {
            return 1;
        }

        if (三级标题编号模板.IsMatch(normalized) && levelIndex == 2)
        {
            return 3;
        }

        if (三级标题编号模板.IsMatch(normalized)
            && levelIndex == 0
            && numId != null
            && !编号实例绑定到非标题样式(mainPart, numId.Value))
        {
            return 3;
        }

        return 0;
    }

    // 目录段落判定（TOC*/目录* 样式）供排版侧与门禁侧共用，避免两处口径漂移
    internal static bool 是目录段落(Paragraph paragraph)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        return !string.IsNullOrWhiteSpace(styleId)
            && (styleId.StartsWith("TOC", StringComparison.OrdinalIgnoreCase)
                || styleId.StartsWith("目录", StringComparison.Ordinal));
    }

    public static int ResolveHeadingLevelForValidation(MainDocumentPart? mainPart, Paragraph p)
    {
        var text = NormalizeText(ParagraphText(p));
        if (text.Length <= 30)
        {
            if (一级标题文本.IsMatch(text)) return 1;
            if (二级标题文本.IsMatch(text)) return 2;
            if (三级标题文本.IsMatch(text)) return 3;
        }

        var visibleNumberingLevel = ResolveHeadingLevelByVisibleNumbering(mainPart, p);
        if (visibleNumberingLevel != 0) return visibleNumberingLevel;

        if (text.Length > 30) return 0;

        var outlineLevel = ResolveOutlineLevel(mainPart, p);
        if (outlineLevel == 0) return 1;
        if (outlineLevel == 1) return 2;
        if (outlineLevel == 2) return 3;

        return ResolveHeadingLevelByStyleId(p.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
    }

    public static int ResolveHeadingLevelByStyleId(string? styleVal)
    {
        if (string.IsNullOrWhiteSpace(styleVal)) return 0;
        var key = styleVal.Trim().ToLowerInvariant();
        if (key is "1" or "heading1" or "标题1" or "biaoti1" or "h1") return 1;
        if (key is "2" or "heading2" or "标题2" or "biaoti2" or "h2") return 2;
        if (key is "3" or "heading3" or "标题3" or "biaoti3" or "h3") return 3;
        return 0;
    }

    public static bool 是正文式四级(MainDocumentPart? mainPart, Paragraph p)
    {
        var text = NormalizeText(ParagraphText(p));
        if (!正文式四级文本.IsMatch(text)) return false;
        return ResolveHeadingLevelForValidation(mainPart, p) == 0;
    }

    public static string Resolve编号来源(MainDocumentPart? mainPart, Paragraph p)
    {
        var direct = p.ParagraphProperties?.GetFirstChild<NumberingProperties>();
        if (direct != null)
        {
            return "段落直接编号";
        }

        var styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var styleNumbering = ResolveStyleNumberingProperties(mainPart, styleId);
        if (styleNumbering != null)
        {
            return "样式链编号";
        }

        return "无";
    }

    public static string 提取可见文本(OpenXmlElement? element)
    {
        if (element == null) return string.Empty;

        var sb = new StringBuilder();
        追加可见文本(element, sb);
        return sb.ToString();
    }

    public static string 提取文本框文本(OpenXmlElement? element)
    {
        if (element == null) return string.Empty;

        var sb = new StringBuilder();
        追加文本框文本(element, sb, false);
        if (sb.Length > 0)
        {
            return sb.ToString();
        }

        return 从未知结构中提取文本框文本(element.OuterXml);
    }

    public static void 物化编号到段落(MainDocumentPart? mainPart, Paragraph p)
    {
        EnsureParagraphProperties(p);
        var pPr = p.ParagraphProperties!;
        if (pPr.GetFirstChild<NumberingProperties>() != null)
        {
            return;
        }

        var styleId = pPr.ParagraphStyleId?.Val?.Value;
        var styleNumbering = ResolveStyleNumberingProperties(mainPart, styleId);
        if (styleNumbering == null)
        {
            return;
        }

        // numPr 在 CT_PPr schema 中必须排在 pStyle/keepNext/keepLines/pageBreakBefore/framePr/widowControl 之后
        // 不能裸 PrependChild，否则会把 numPr 放到这些元素前面导致 schema 违规
        InsertAfterPrecedingElements(pPr, styleNumbering,
            ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl"],
            ["suppressLineNumbers", "pBdr", "shd", "tabs"]);
    }

    /// <summary>
    /// 在 parent 中按 schema 顺序插入 node：先找 precedingLocalNames 中最后一个已有元素并 InsertAfter；
    /// 若都不存在，则找 followingLocalNames 中第一个已有元素并 InsertBefore；
    /// 若仍无锚点，插到最前面（此时 parent 中没有排在 node 之前的元素，追加到末尾反而会落到 schema 顺序更靠后的元素之后）。
    /// </summary>
    private static void InsertAfterPrecedingElements(
        OpenXmlElement parent, OpenXmlElement node,
        string[] precedingLocalNames, string[] followingLocalNames)
    {
        // 先倒序找最后一个 "应在 node 之前" 的已有子元素
        for (var i = precedingLocalNames.Length - 1; i >= 0; i--)
        {
            var anchor = parent.ChildElements
                .FirstOrDefault(c => c.LocalName == precedingLocalNames[i]);
            if (anchor != null)
            {
                parent.InsertAfter(node, anchor);
                return;
            }
        }

        // 再正序找第一个 "应在 node 之后" 的已有子元素
        foreach (var name in followingLocalNames)
        {
            var anchor = parent.ChildElements
                .FirstOrDefault(c => c.LocalName == name);
            if (anchor != null)
            {
                parent.InsertBefore(node, anchor);
                return;
            }
        }

        // 无锚点：说明 parent 里没有排在 node 之前的元素，应插到最前面而非追加末尾
        if (parent.FirstChild == null)
        {
            parent.AppendChild(node);
        }
        else
        {
            parent.InsertBefore(node, parent.FirstChild);
        }
    }

    public static bool 段落存在编号(MainDocumentPart? mainPart, Paragraph p)
    {
        return ResolveParagraphNumberingProperties(mainPart, p) != null;
    }

    public static int? ResolveOutlineLevel(MainDocumentPart? mainPart, Paragraph p)
    {
        var direct = p.ParagraphProperties?.OutlineLevel?.Val?.Value;
        if (direct.HasValue)
        {
            return direct.Value;
        }

        var styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (string.IsNullOrWhiteSpace(styleId))
        {
            return null;
        }

        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var currentStyleId = styleId;
        while (!string.IsNullOrWhiteSpace(currentStyleId) && visited.Add(currentStyleId))
        {
            var style = FindStyle(mainPart, StyleValues.Paragraph, currentStyleId);
            var outline = style?.StyleParagraphProperties?.OutlineLevel?.Val?.Value;
            if (outline.HasValue)
            {
                return outline.Value;
            }

            currentStyleId = style?.BasedOn?.Val?.Value;
        }

        return null;
    }

    public static void SetRunFonts(RunProperties runProps, string eastAsia, string latin)
    {
        runProps.RunFonts ??= new RunFonts();
        runProps.RunFonts.Ascii = latin;
        runProps.RunFonts.HighAnsi = latin;
        runProps.RunFonts.ComplexScript = latin;
        runProps.RunFonts.EastAsia = eastAsia;
    }

    public static void EnsureParagraphProperties(Paragraph p)
    {
        if (p.ParagraphProperties is null)
        {
            p.PrependChild(new ParagraphProperties());
        }
    }

    // 只处理本段的运行块，不能从承载图片的外层段落穿透到文本框内部。
    public static IEnumerable<Run> Runs(Paragraph p) => p.Descendants<Run>()
        .Where(run => ReferenceEquals(run.Ancestors<Paragraph>().FirstOrDefault(), p));

    internal static bool 是隐藏运行块(Run run) =>
        (run.RunProperties?.Vanish is { } vanish && (vanish.Val?.Value ?? true))
        || (run.RunProperties?.WebHidden is { } webHidden && (webHidden.Val?.Value ?? true));

    public static bool IsStyleBold(MainDocumentPart? mainPart, StyleValues styleType, string? styleId)
    {
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var style = FindStyle(mainPart, styleType, styleId);
        while (style != null && visited.Add(style.StyleId?.Value ?? string.Empty))
        {
            var bold = style.StyleRunProperties?.Bold;
            if (bold != null) return bold.Val == null || bold.Val.Value;
            var nextStyleId = style.BasedOn?.Val?.Value;
            style = FindStyle(mainPart, styleType, nextStyleId);
        }
        return false;
    }

    public static int? GetStyleFontSize(MainDocumentPart? mainPart, StyleValues styleType, string? styleId)
    {
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var style = FindStyle(mainPart, styleType, styleId);
        while (style != null && visited.Add(style.StyleId?.Value ?? string.Empty))
        {
            var value = style.StyleRunProperties?.FontSize?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(value) && int.TryParse(value, out var size))
            {
                return size;
            }
            var nextStyleId = style.BasedOn?.Val?.Value;
            style = FindStyle(mainPart, styleType, nextStyleId);
        }
        return null;
    }

    private static void 追加可见文本(OpenXmlElement element, StringBuilder sb, Paragraph? owner = null)
    {
        if (owner != null && element is Paragraph paragraph && !ReferenceEquals(paragraph, owner)) return;
        if (element is Drawing || element.LocalName is "anchor" or "inline")
        {
            return;
        }

        if (element is Run run)
        {
            if (是隐藏运行块(run))
            {
                return;
            }
        }

        switch (element)
        {
            case Text text:
                sb.Append(text.Text);
                return;
            case TabChar:
                sb.Append('\t');
                return;
            case Break:
            case CarriageReturn:
                sb.Append('\n');
                return;
        }

        if (element.LocalName == "delText" || element.LocalName == "instrText")
        {
            return;
        }

        foreach (var child in element.ChildElements)
        {
            追加可见文本(child, sb, owner);
        }
    }

    private static void 追加文本框文本(OpenXmlElement element, StringBuilder sb, bool 位于文本框内)
    {
        var nowInside = 位于文本框内 || element is TextBoxContent || element.LocalName == "txbxContent";

        if (nowInside)
        {
            switch (element)
            {
                case Text text:
                    sb.Append(text.Text);
                    return;
                case TabChar:
                    sb.Append('\t');
                    return;
                case Break:
                case CarriageReturn:
                    sb.Append('\n');
                    return;
            }
        }

        foreach (var child in element.ChildElements)
        {
            追加文本框文本(child, sb, nowInside);
        }
    }

    private static string 从未知结构中提取文本框文本(string? xml)
    {
        if (string.IsNullOrWhiteSpace(xml))
        {
            return string.Empty;
        }

        var result = new StringBuilder();
        foreach (Match block in 文本框片段.Matches(xml))
        {
            foreach (Match text in 文本框文字节点.Matches(block.Value))
            {
                result.Append(System.Net.WebUtility.HtmlDecode(text.Groups[1].Value));
            }
        }

        return result.ToString();
    }

    private static Style? FindStyle(MainDocumentPart? mainPart, StyleValues styleType, string? styleId)
    {
        if (mainPart?.StyleDefinitionsPart?.Styles == null || string.IsNullOrWhiteSpace(styleId))
        {
            return null;
        }
        return mainPart.StyleDefinitionsPart.Styles.Elements<Style>()
            .FirstOrDefault(s => s.StyleId?.Value == styleId && s.Type?.Value == styleType);
    }

    public static Level? ResolveNumberingLevel(MainDocumentPart? mainPart, Paragraph p)
    {
        var numPr = ResolveParagraphNumberingProperties(mainPart, p);
        var numId = numPr?.GetFirstChild<NumberingId>()?.Val?.Value;
        var levelIndex = numPr?.GetFirstChild<NumberingLevelReference>()?.Val?.Value;
        if (mainPart?.NumberingDefinitionsPart?.Numbering == null || numId == null || levelIndex == null)
        {
            return null;
        }

        var numbering = mainPart.NumberingDefinitionsPart.Numbering;
        var num = numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        var abstractNumId = num?.AbstractNumId?.Val?.Value;
        if (abstractNumId == null)
        {
            return null;
        }

        var abstractNum = numbering.Elements<AbstractNum>()
            .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId);
        if (abstractNum == null)
        {
            return null;
        }

        return abstractNum.Elements<Level>()
            .FirstOrDefault(l => l.LevelIndex?.Value == levelIndex);
    }

    public static NumberingProperties? ResolveParagraphNumberingProperties(MainDocumentPart? mainPart, Paragraph p)
    {
        var direct = p.ParagraphProperties?.GetFirstChild<NumberingProperties>();
        if (direct != null)
        {
            return direct;
        }

        var styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        return ResolveStyleNumberingProperties(mainPart, styleId);
    }

    private static NumberingProperties? ResolveStyleNumberingProperties(MainDocumentPart? mainPart, string? styleId)
    {
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var currentStyleId = styleId;

        while (!string.IsNullOrWhiteSpace(currentStyleId) && visited.Add(currentStyleId))
        {
            var style = FindStyle(mainPart, StyleValues.Paragraph, currentStyleId);
            var numPr = style?.StyleParagraphProperties?.GetFirstChild<NumberingProperties>();
            if (numPr != null)
            {
                return numPr.CloneNode(true) as NumberingProperties;
            }

            currentStyleId = style?.BasedOn?.Val?.Value;
        }

        return null;
    }

    private static bool 编号实例绑定到非标题样式(MainDocumentPart? mainPart, int numId)
    {
        var styles = mainPart?.StyleDefinitionsPart?.Styles;
        if (styles == null) return false;

        foreach (var style in styles.Elements<Style>())
        {
            if (style.Type != null && style.Type.Value != StyleValues.Paragraph) continue;

            var styleId = style.StyleId?.Value;
            if (string.IsNullOrWhiteSpace(styleId)) continue;

            var styleNumPr = ResolveStyleNumberingProperties(mainPart, styleId);
            var styleNumId = styleNumPr?.GetFirstChild<NumberingId>()?.Val?.Value;
            if (styleNumId != numId) continue;

            if (ResolveHeadingLevelByStyleId(styleId) == 0)
            {
                return true;
            }
        }

        return false;
    }

    private static bool 是括号汉字二级标题编号(Level level, string normalizedLevelText)
    {
        if (!二级标题编号模板.IsMatch(normalizedLevelText))
        {
            return false;
        }

        var numberingFormat = level.GetFirstChild<NumberingFormat>()?.Val?.Value;
        return numberingFormat != null && numberingFormat != NumberFormatValues.Decimal;
    }

    /// <summary>
    /// 清理旧版 Word 兼容格式遗留在 SectionProperties 中的非法元素。
    /// 例如 evenAndOddHeaders 属于 w:settings，不应出现在 w:sectPr 中。
    /// </summary>
    public static void CleanupLegacySectionProperties(WordprocessingDocument word)
    {
        var mainPart = word.MainDocumentPart;
        var body = mainPart?.Document?.Body;
        if (mainPart == null || body == null) return;

        // 收集所有 SectionProperties（包括段落内分节和 Body 直属）
        var sections = new List<SectionProperties>();
        foreach (var p in body.Elements<Paragraph>())
        {
            var sectPr = p.ParagraphProperties?.GetFirstChild<SectionProperties>();
            if (sectPr != null) sections.Add(sectPr);
        }
        var bodySect = body.GetFirstChild<SectionProperties>();
        if (bodySect != null) sections.Add(bodySect);

        // 也扫描页眉/页脚 XML 部件中的 SectionProperties（旧版 Word 可能把 EvenAndOddHeaders 也误放到这里）
        foreach (var headerPart in mainPart.HeaderParts)
        {
            var sectPr = headerPart.Header?.GetFirstChild<SectionProperties>();
            if (sectPr != null) sections.Add(sectPr);
        }
        foreach (var footerPart in mainPart.FooterParts)
        {
            var sectPr = footerPart.Footer?.GetFirstChild<SectionProperties>();
            if (sectPr != null) sections.Add(sectPr);
        }

        foreach (var sectPr in sections.Distinct())
        {
            // evenAndOddHeaders 属于 w:settings，旧版 Word 可能误放到 w:sectPr
            var evenAndOdd = sectPr.GetFirstChild<EvenAndOddHeaders>();
            evenAndOdd?.Remove();

            // titlePg 是 w:sectPr 的合法子元素，但旧版 Word 可能在封面节之后仍保留
            // 这里不强制移除 titlePg，由 HeaderFooterService 按需处理
        }
    }

    /// <summary>
    /// 按 ECMA-376 schema 规定的子元素顺序重排 w:style 的直接子元素。
    /// w:style 子元素顺序：name → aliases → basedOn → next → link → autoRedefine →
    /// hidden → uiPriority → semiHidden → unhideWhenUsed → qFormat → locked →
    /// personal → personalCompose → personalReply → rsid → pPr → rPr → tblPr → trPr → tcPr → tblStylePr
    /// </summary>
    public static void NormalizeStyleChildOrder(WordprocessingDocument word)
    {
        var styles = word.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles == null) return;

        // CT_Style 子元素按 schema 规定顺序排列（用 LocalName 匹配）
        var schemaOrder = new[]
        {
            "name", "aliases", "basedOn", "next", "link", "autoRedefine",
            "hidden", "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked",
            "personal", "personalCompose", "personalReply", "rsid",
            "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr"
        };

        foreach (var style in styles.Elements<Style>())
        {
            ReorderChildren(style, schemaOrder);
        }
    }

    // 子元素顺序见 ECMA-376 CT_RPr；rPrChange 排在最后
    private static readonly string[] RunPropertiesSchemaOrder =
    [
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
        "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
        "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
        "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
        "specVanish", "oMath", "rPrChange"
    ];

    // 子元素顺序见 ECMA-376 CT_TcPr；tcPrChange 排在最后
    private static readonly string[] TableCellPropertiesSchemaOrder =
    [
        "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap",
        "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers",
        "cellIns", "cellDel", "cellMerge", "tcPrChange"
    ];

    /// <summary>
    /// 按 schema 顺序重排正文与页眉页脚中的 run 属性(rPr)、单元格属性(tcPr)子元素，
    /// 并给缺失 val 的底纹补齐默认值。
    /// 输入文档常由脚本生成，元素顺序不合规会被 OpenXmlValidator 报为 unexpected child element；
    /// 重排只调整顺序、补默认属性，元素集合与显示效果均不变。
    /// </summary>
    public static void NormalizeRunAndCellPropertiesOrder(WordprocessingDocument word)
    {
        var mainPart = word.MainDocumentPart;
        if (mainPart == null) return;

        var roots = new List<OpenXmlElement>();
        if (mainPart.Document != null) roots.Add(mainPart.Document);
        foreach (var headerPart in mainPart.HeaderParts)
        {
            if (headerPart.Header != null) roots.Add(headerPart.Header);
        }
        foreach (var footerPart in mainPart.FooterParts)
        {
            if (footerPart.Footer != null) roots.Add(footerPart.Footer);
        }

        foreach (var root in roots)
        {
            foreach (var runProperties in root.Descendants<RunProperties>())
            {
                ReorderChildrenIfAllKnown(runProperties, RunPropertiesSchemaOrder);
            }

            foreach (var cellProperties in root.Descendants<TableCellProperties>())
            {
                ReorderChildrenIfAllKnown(cellProperties, TableCellPropertiesSchemaOrder);
                foreach (var shading in cellProperties.Elements<Shading>())
                {
                    // CT_Shd 的 val 是必需属性，缺失时按“清除图案、纯色填充”补齐
                    shading.Val ??= ShadingPatternValues.Clear;
                }
            }
        }
    }

    private static void ReorderChildrenIfAllKnown(OpenXmlElement parent, string[] schemaOrder)
    {
        // 出现顺序表之外的子元素时保守跳过，否则会被排到末尾，反而制造新的 schema 违规
        foreach (var child in parent.ChildElements)
        {
            if (!schemaOrder.Contains(child.LocalName, StringComparer.OrdinalIgnoreCase)) return;
        }
        ReorderChildren(parent, schemaOrder);
    }

    private static void ReorderChildren(OpenXmlElement parent, string[] schemaOrder)
    {
        var children = parent.ChildElements.ToList();
        if (children.Count <= 1) return;

        // 建立 schema 顺序映射：LocalName → 排序权重
        var orderMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < schemaOrder.Length; i++)
        {
            orderMap[schemaOrder[i]] = i;
        }

        // 检查是否需要重排（当前顺序是否已正确）
        var needsReorder = false;
        var lastOrder = -1;
        foreach (var child in children)
        {
            if (!orderMap.TryGetValue(child.LocalName, out var currentOrder))
            {
                currentOrder = int.MaxValue; // 未知元素放最后
            }
            if (currentOrder < lastOrder)
            {
                needsReorder = true;
                break;
            }
            lastOrder = currentOrder;
        }

        if (!needsReorder) return;

        // 按 schema 顺序排序并重建
        var sorted = children
            .OrderBy(c => orderMap.TryGetValue(c.LocalName, out var o) ? o : int.MaxValue)
            .ToList();

        parent.RemoveAllChildren();
        foreach (var child in sorted)
        {
            parent.AppendChild(child);
        }
    }
}
