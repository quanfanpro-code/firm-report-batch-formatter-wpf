using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FirmFormatter.OpenXml.Core;

// CT_SectPr 子元素必须按 ECMA-376 schema 顺序排列：
// headerReference → footerReference → footnotePr → endnotePr → type →
// pgSz → pgMar → paperSrc → pgBorders → lnNumType → pgNumType →
// cols → formProt → vAlign → noEndnote → titlePg → textDirection →
// bidi → rtlGutter → docGrid → printerSettings → sectPrChange
// 新增子元素时务必用 Ensure* 方法插入到 schema 合法位置，禁止裸 AppendChild。

public sealed class HeaderFooterService
{
    private const uint PortraitWidth = 11906U;
    private const uint PortraitHeight = 16838U;
    private const uint LandscapeWidth = 16838U;
    private const uint LandscapeHeight = 11906U;

    private readonly FirmRuleProfile _ruleProfile = FirmRuleProfile.Default;

    public void Apply(WordprocessingDocument word, bool hasCover, bool isPureCover = false, CancellationToken cancellationToken = default)
    {
        var main = word.MainDocumentPart;
        var body = main?.Document?.Body;
        if (main is null || body is null) return;

        EnsureFinalSection(body);
        var sections = OpenXmlHelper.收集分节(body);
        var hasDedicatedCoverSection = hasCover && sections.Count > 1;
        var evenAndOdd = main.DocumentSettingsPart?.Settings?.GetFirstChild<EvenAndOddHeaders>();
        var useEvenPageFooter = evenAndOdd is not null && (evenAndOdd.Val?.Value ?? true);
        FooterPart? sharedFooterPart = null;
        for (var i = 0; i < sections.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sectPr = sections[i];
            var treatAsCover = isPureCover || (hasDedicatedCoverSection && i == 0);
            EnsurePageLayout(sectPr, treatAsCover, _ruleProfile);

            if (treatAsCover)
            {
                ClearHeaderFooter(main, sectPr);
                continue;
            }

            if (hasDedicatedCoverSection && i == 1)
            {
                var pgNum = EnsurePageNumberType(sectPr);
                pgNum.Start = 1;
            }

            FormatExistingHeaders(main, sectPr, _ruleProfile);
            ReplaceFooterWithPageField(main, sectPr, _ruleProfile, useEvenPageFooter, ref sharedFooterPart);
        }
    }

    public void ApplyPureCoverLayout(WordprocessingDocument word)
    {
        var body = word.MainDocumentPart?.Document?.Body;
        if (body is null) return;

        EnsureFinalSection(body);
        foreach (var sectPr in OpenXmlHelper.收集分节(body))
        {
            EnsurePureCoverMargins(sectPr, _ruleProfile);
        }
    }

    private static void EnsureFinalSection(Body body)
    {
        // 末节属性在输入中可以省略，但输出必须显式写入，才能实际设置纸张和页边距。
        if (!body.Elements<SectionProperties>().Any()) body.AppendChild(new SectionProperties());
    }

    private static void EnsurePageLayout(SectionProperties sectPr, bool isCover, FirmRuleProfile ruleProfile)
    {
        var pgSz = EnsurePageSize(sectPr);
        var pgMar = EnsurePageMargin(sectPr);

        pgMar.Top = ruleProfile.页边距上Twips;
        pgMar.Left = (UInt32Value)(uint)ruleProfile.页边距左Twips;
        pgMar.Right = (UInt32Value)(uint)ruleProfile.页边距右Twips;
        pgMar.Bottom = isCover ? ruleProfile.纯封面下边距Twips : ruleProfile.普通文档下边距Twips;
        pgMar.Header = (UInt32Value)ruleProfile.页眉距离Twips;
        pgMar.Footer = (UInt32Value)ruleProfile.页脚距离Twips;

        var isLandscape = IsLandscapeSection(pgSz);
        if (isLandscape)
        {
            pgSz.Width = LandscapeWidth;
            pgSz.Height = LandscapeHeight;
            pgSz.Orient = PageOrientationValues.Landscape;
            return;
        }

        pgSz.Width = PortraitWidth;
        pgSz.Height = PortraitHeight;
        pgSz.Orient = PageOrientationValues.Portrait;
    }

    private static void EnsurePureCoverMargins(SectionProperties sectPr, FirmRuleProfile ruleProfile)
    {
        var pgMar = EnsurePageMargin(sectPr);
        pgMar.Top = ruleProfile.页边距上Twips;
        pgMar.Left = (UInt32Value)(uint)ruleProfile.页边距左Twips;
        pgMar.Right = (UInt32Value)(uint)ruleProfile.页边距右Twips;
        pgMar.Bottom = ruleProfile.纯封面下边距Twips;
        pgMar.Header = (UInt32Value)ruleProfile.页眉距离Twips;
        pgMar.Footer = (UInt32Value)ruleProfile.页脚距离Twips;
    }

    private static PageSize EnsurePageSize(SectionProperties sectPr)
    {
        var existing = sectPr.GetFirstChild<PageSize>();
        if (existing != null) return existing;

        var node = new PageSize();
        var anchor = FindAnchorAfter(sectPr, SchemaGroup.AfterPgSz);
        InsertAtAnchor(sectPr, node, anchor);
        return node;
    }

    private static PageMargin EnsurePageMargin(SectionProperties sectPr)
    {
        var existing = sectPr.GetFirstChild<PageMargin>();
        if (existing != null) return existing;

        var node = new PageMargin();
        var anchor = FindAnchorAfter(sectPr, SchemaGroup.AfterPgMar);
        InsertAtAnchor(sectPr, node, anchor);
        return node;
    }

    private static PageNumberType EnsurePageNumberType(SectionProperties sectPr)
    {
        var existing = sectPr.GetFirstChild<PageNumberType>();
        if (existing != null) return existing;

        var node = new PageNumberType();
        var anchor = FindAnchorAfter(sectPr, SchemaGroup.AfterPgNumType);
        InsertAtAnchor(sectPr, node, anchor);
        return node;
    }

    private static OpenXmlElement? FindAnchorAfter(SectionProperties sectPr, string[] laterLocalNames)
    {
        foreach (var child in sectPr.ChildElements)
        {
            if (Array.Exists(laterLocalNames, n => child.LocalName == n))
                return child;
        }
        return null;
    }

    private static void InsertAtAnchor(SectionProperties sectPr, OpenXmlElement node, OpenXmlElement? anchor)
    {
        if (anchor != null)
            sectPr.InsertBefore(node, anchor);
        else
            sectPr.AppendChild(node);
    }

    private static class SchemaGroup
    {
        internal static readonly string[] AfterPgSz =
            ["pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType",
             "cols", "formProt", "vAlign", "noEndnote", "titlePg",
             "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"];

        internal static readonly string[] AfterPgMar =
            ["paperSrc", "pgBorders", "lnNumType", "pgNumType",
             "cols", "formProt", "vAlign", "noEndnote", "titlePg",
             "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"];

        internal static readonly string[] AfterPgNumType =
            ["cols", "formProt", "vAlign", "noEndnote", "titlePg",
             "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"];
    }

    private static bool IsLandscapeSection(PageSize pgSz)
    {
        if (pgSz.Orient?.Value == PageOrientationValues.Landscape)
        {
            return true;
        }

        var width = pgSz.Width?.Value;
        var height = pgSz.Height?.Value;
        return width.HasValue && height.HasValue && width.Value > height.Value;
    }

    private static void ClearHeaderFooter(MainDocumentPart main, SectionProperties sectPr)
    {
        foreach (var headerReference in sectPr.Elements<HeaderReference>().ToList())
        {
            移除页眉引用(main, sectPr, headerReference);
        }

        foreach (var footerReference in sectPr.Elements<FooterReference>().ToList())
        {
            移除页脚引用(main, sectPr, footerReference);
        }

        var titlePg = sectPr.GetFirstChild<TitlePage>();
        titlePg?.Remove();
    }

    private static void FormatExistingHeaders(MainDocumentPart main, SectionProperties sectPr, FirmRuleProfile ruleProfile)
    {
        // 版心宽 = 页宽 − 左右边距（EnsurePageLayout 已兜底写值），横版分节同样适用
        var pgSz = sectPr.GetFirstChild<PageSize>();
        var pgMar = sectPr.GetFirstChild<PageMargin>();
        var contentWidthTwips = (int)((pgSz?.Width?.Value ?? PortraitWidth)
            - (pgMar?.Left?.Value ?? (uint)ruleProfile.页边距左Twips)
            - (pgMar?.Right?.Value ?? (uint)ruleProfile.页边距右Twips));

        foreach (var hr in sectPr.Elements<HeaderReference>().ToList())
        {
            // 死引用（缺关系 Id 或指向不存在的部件）既无内容可排版，又会触发门禁
            // header_ref_missing_id/header_ref_broken 拦截，甚至让 OpenXmlValidator 崩溃，直接清理
            if (string.IsNullOrWhiteSpace(hr.Id?.Value))
            {
                hr.Remove();
                continue;
            }

            HeaderPart? headerPart;
            // 仅容忍关系 ID 无效这一预期异常，其余异常应暴露而不是静默跳过
            try { headerPart = main.GetPartById(hr.Id.Value) as HeaderPart; }
            catch (ArgumentOutOfRangeException) { hr.Remove(); continue; }
            if (headerPart is null) { hr.Remove(); continue; }
            // 横竖版不能共用一个绝对定位点；仅在目标宽度不同时拆开页眉并保留全部关系。
            var landscape = IsLandscapeSection(sectPr.GetFirstChild<PageSize>() ?? new PageSize());
            if (OpenXmlHelper.收集分节(main.Document?.Body).Any(section =>
                !ReferenceEquals(section, sectPr)
                && IsLandscapeSection(section.GetFirstChild<PageSize>() ?? new PageSize()) != landscape
                && section.Elements<HeaderReference>().Any(reference => reference.Id?.Value == hr.Id.Value)))
            {
                var copy = main.AddNewPart<HeaderPart>();
                copy.Header = headerPart.Header?.CloneNode(true) as Header;
                foreach (var part in headerPart.Parts) copy.AddPart(part.OpenXmlPart, part.RelationshipId);
                foreach (var link in headerPart.HyperlinkRelationships) copy.AddHyperlinkRelationship(link.Uri, link.IsExternal, link.Id);
                foreach (var link in headerPart.ExternalRelationships) copy.AddExternalRelationship(link.RelationshipType, link.Uri, link.Id);
                foreach (var link in headerPart.DataPartReferenceRelationships)
                    copy.AddVideoReferenceRelationship((MediaDataPart)link.DataPart, link.Id);
                hr.Id = main.GetIdOfPart(copy);
                headerPart = copy;
            }
            var header = headerPart.Header;
            if (header is null) continue;

            foreach (var p in header.Descendants<Paragraph>().Where(p => !p.Ancestors<TextBoxContent>().Any()))
            {
                var originalTabs = p.ParagraphProperties?.Tabs?.CloneNode(true) as Tabs;
                ResetHeaderParagraphProperties(p);
                // 已有制表符不再清掉定位点，二次排版及共享页眉保持左右对齐。
                if (p.Descendants<TabChar>().Any())
                {
                    p.ParagraphProperties!.Tabs = originalTabs;
                    foreach (var tab in originalTabs?.Elements<TabStop>() ?? [])
                        if (tab.Val?.Value == TabStopValues.Right) tab.Position = contentWidthTwips;
                }
                规范化页眉内容(p, contentWidthTwips);

                foreach (var run in OpenXmlHelper.Runs(p))
                {
                    var vanish = run.RunProperties?.Vanish?.CloneNode(true) as Vanish;
                    var webHidden = run.RunProperties?.WebHidden?.CloneNode(true) as WebHidden;
                    run.RunProperties = CreateHeaderRunProperties(ruleProfile);
                    run.RunProperties.Vanish = vanish;
                    run.RunProperties.WebHidden = webHidden;
                }
            }
            header.Save();
        }
    }

    // "文字 + 连续 2 个及以上空格（含全角）+ 文字"是用空格凑右对齐的手工排版，
    // 空格宽度随字号变化，统一字号后必然错位，必须换成右对齐定位点
    private static readonly System.Text.RegularExpressions.Regex 页眉填充正则 =
        new(@"^(?<left>\S(?:.*?\S)?)[ 　]{2,}(?<right>\S.*)$", System.Text.RegularExpressions.RegexOptions.Compiled);

    private static void 规范化页眉内容(Paragraph paragraph, int contentWidthTwips)
    {
        if (OpenXmlHelper.Runs(paragraph).Any(run => run.ChildElements.Any(child => child is not RunProperties and not Text and not TabChar))) return;
        // 含域、图片、书签、修订等复杂结构的段落不重建，避免破坏内容
        if (paragraph.Descendants().Any(element =>
                element is Drawing or SimpleField or FieldCode or Hyperlink
                    or BookmarkStart or BookmarkEnd or InsertedRun or DeletedRun
                    or Break or Vanish or WebHidden or FootnoteReference or EndnoteReference
                || element.LocalName is "sdt" or "oMath" or "oMathPara" or "object" or "pict"
                    or "commentRangeStart" or "commentRangeEnd" or "commentReference"))
        {
            return;
        }

        var fullText = string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));
        var trimmed = fullText.TrimStart(' ', '　');

        // 行首制表符 = 第一个非空白文字之前出现的 TabChar（规范要求页眉左对齐顶格）
        var hasLeadingTab = false;
        foreach (var element in paragraph.Descendants())
        {
            if (element is TabChar) hasLeadingTab = true;
            else if (element is Text t && t.Text.Any(c => c is not ' ' and not '　')) break;
        }

        var match = 页眉填充正则.Match(trimmed);
        if (!match.Success && !hasLeadingTab && trimmed.Length == fullText.Length) return;

        // 重建内容（守卫已排除复杂结构，这里只剩普通 run）
        foreach (var child in paragraph.ChildElements.Where(c => c is not ParagraphProperties).ToList())
        {
            child.Remove();
        }

        if (match.Success)
        {
            paragraph.AppendChild(new Run(new Text(match.Groups["left"].Value) { Space = SpaceProcessingModeValues.Preserve }));
            paragraph.AppendChild(new Run(new TabChar(), new Text(match.Groups["right"].Value.TrimEnd(' ', '　')) { Space = SpaceProcessingModeValues.Preserve }));
            paragraph.ParagraphProperties!.Tabs = new Tabs(new TabStop
            {
                Val = TabStopValues.Right,
                Position = contentWidthTwips
            });
        }
        else
        {
            paragraph.AppendChild(new Run(new Text(trimmed) { Space = SpaceProcessingModeValues.Preserve }));
        }
    }

    private static void ResetHeaderParagraphProperties(Paragraph paragraph)
    {
        var sectionProperties = paragraph.ParagraphProperties?.GetFirstChild<SectionProperties>()?.CloneNode(true);
        var pPr = new ParagraphProperties();
        if (sectionProperties != null)
        {
            pPr.Append(sectionProperties);
        }

        pPr.Justification = new Justification { Val = JustificationValues.Left };
        pPr.Indentation = new Indentation
        {
            Left = "0",
            Right = "0",
            Start = "0",
            FirstLine = "0",
            LeftChars = 0,
            RightChars = 0,
            StartCharacters = 0,
            FirstLineChars = 0
        };
        pPr.SpacingBetweenLines = new SpacingBetweenLines
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
        pPr.ParagraphBorders = new ParagraphBorders(
            new BottomBorder
            {
                Val = BorderValues.Single,
                Size = 6U,
                Space = 1U,
                Color = "auto"
            });
        paragraph.ParagraphProperties = pPr;
    }

    private static RunProperties CreateHeaderRunProperties(FirmRuleProfile ruleProfile)
    {
        var runProperties = new RunProperties
        {
            FontSize = new FontSize { Val = ruleProfile.页眉页脚字号HalfPoint },
            FontSizeComplexScript = new FontSizeComplexScript { Val = ruleProfile.页眉页脚字号HalfPoint },
            Bold = new Bold { Val = false },
            BoldComplexScript = new BoldComplexScript { Val = false },
            Italic = new Italic { Val = false },
            ItalicComplexScript = new ItalicComplexScript { Val = false }
        };
        OpenXmlHelper.SetRunFonts(runProperties, ruleProfile.中文字体, ruleProfile.西文字体);
        return runProperties;
    }

    private static void ReplaceFooterWithPageField(MainDocumentPart main, SectionProperties sectPr, FirmRuleProfile ruleProfile, bool useEvenPageFooter, ref FooterPart? sharedFooterPart)
    {
        foreach (var fr in sectPr.Elements<FooterReference>().ToList())
        {
            移除页脚引用(main, sectPr, fr);
        }

        // 所有非封面分节共用同一个页脚部件，避免每个分节重复创建内容完全相同的 FooterPart
        sharedFooterPart ??= CreatePageFieldFooterPart(main, ruleProfile);
        var relId = main.GetIdOfPart(sharedFooterPart);
        插入页脚引用到合法位置(sectPr, new FooterReference { Id = relId, Type = HeaderFooterValues.Default });

        if (useEvenPageFooter)
        {
            插入页脚引用到合法位置(sectPr, new FooterReference { Id = relId, Type = HeaderFooterValues.Even });
        }

        // 保留"首页不同"语义：该分节若带 titlePg，需同时挂 First 类型页脚引用，
        // 否则该节首页没有任何页脚，页码会在首页断号
        if (sectPr.GetFirstChild<TitlePage>() is not null)
        {
            插入页脚引用到合法位置(sectPr, new FooterReference { Id = relId, Type = HeaderFooterValues.First });
        }
    }

    private static FooterPart CreatePageFieldFooterPart(MainDocumentPart main, FirmRuleProfile ruleProfile)
    {
        var footerPart = main.AddNewPart<FooterPart>();
        var p = new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }));

        var runProp = new RunProperties();
        runProp.FontSize = new FontSize { Val = ruleProfile.页眉页脚字号HalfPoint };
        runProp.FontSizeComplexScript = new FontSizeComplexScript { Val = ruleProfile.页眉页脚字号HalfPoint };
        OpenXmlHelper.SetRunFonts(runProp, ruleProfile.中文字体, ruleProfile.西文字体);

        var runBegin = new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }) { RunProperties = (RunProperties)runProp.CloneNode(true) };
        var runCode = new Run(new FieldCode(" PAGE ")) { RunProperties = (RunProperties)runProp.CloneNode(true) };
        var runSep = new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }) { RunProperties = (RunProperties)runProp.CloneNode(true) };
        var runText = new Run(new Text("1")) { RunProperties = (RunProperties)runProp.CloneNode(true) };
        var runEnd = new Run(new FieldChar { FieldCharType = FieldCharValues.End }) { RunProperties = (RunProperties)runProp.CloneNode(true) };
        p.Append(runBegin, runCode, runSep, runText, runEnd);
        footerPart.Footer = new Footer(p);
        footerPart.Footer.Save();
        return footerPart;
    }

    private static void 插入页脚引用到合法位置(SectionProperties sectPr, FooterReference footerReference)
    {
        var anchor = sectPr.ChildElements
            .FirstOrDefault(x => x is not HeaderReference && x is not FooterReference);
        if (anchor != null)
        {
            sectPr.InsertBefore(footerReference, anchor);
            return;
        }

        sectPr.AppendChild(footerReference);
    }

    private static void 移除页眉引用(MainDocumentPart main, SectionProperties currentSection, HeaderReference headerReference)
    {
        var relationId = headerReference.Id?.Value;
        headerReference.Remove();
        if (string.IsNullOrWhiteSpace(relationId)) return;
        if (仍被其他分节引用(main, currentSection, relationId, true)) return;
        if (main.Parts.Any(part => part.RelationshipId == relationId && part.OpenXmlPart is HeaderPart))
            main.DeletePart(relationId);
    }

    private static void 移除页脚引用(MainDocumentPart main, SectionProperties currentSection, FooterReference footerReference)
    {
        var relationId = footerReference.Id?.Value;
        footerReference.Remove();
        if (string.IsNullOrWhiteSpace(relationId)) return;
        if (仍被其他分节引用(main, currentSection, relationId, false)) return;
        if (main.Parts.Any(part => part.RelationshipId == relationId && part.OpenXmlPart is FooterPart))
            main.DeletePart(relationId);
    }

    private static bool 仍被其他分节引用(MainDocumentPart main, SectionProperties currentSection, string relationId, bool isHeader)
    {
        var body = main.Document?.Body;
        if (body is null) return false;

        foreach (var section in OpenXmlHelper.收集分节(body))
        {
            if (ReferenceEquals(section, currentSection)) continue;

            var stillReferenced = isHeader
                ? section.Elements<HeaderReference>().Any(x => x.Id?.Value == relationId)
                : section.Elements<FooterReference>().Any(x => x.Id?.Value == relationId);
            if (stillReferenced) return true;
        }

        return false;
    }
}
