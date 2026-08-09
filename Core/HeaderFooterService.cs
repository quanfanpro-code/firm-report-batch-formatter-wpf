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

        foreach (var sectPr in OpenXmlHelper.收集分节(body))
        {
            EnsurePureCoverMargins(sectPr, _ruleProfile);
        }
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
        foreach (var hr in sectPr.Elements<HeaderReference>())
        {
            if (hr.Id?.Value is null) continue;
            HeaderPart? headerPart;
            // 仅容忍关系 ID 无效这一预期异常，其余异常应暴露而不是静默跳过
            try { headerPart = main.GetPartById(hr.Id.Value) as HeaderPart; }
            catch (ArgumentOutOfRangeException) { continue; }
            if (headerPart is null) continue;
            var header = headerPart.Header;
            if (header is null) continue;

            foreach (var p in header.Descendants<Paragraph>())
            {
                ResetHeaderParagraphProperties(p);

                foreach (var run in p.Descendants<Run>())
                {
                    run.RunProperties = CreateHeaderRunProperties(ruleProfile);
                }
            }
            header.Save();
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
        main.DeletePart(relationId);
    }

    private static void 移除页脚引用(MainDocumentPart main, SectionProperties currentSection, FooterReference footerReference)
    {
        var relationId = footerReference.Id?.Value;
        footerReference.Remove();
        if (string.IsNullOrWhiteSpace(relationId)) return;
        if (仍被其他分节引用(main, currentSection, relationId, false)) return;
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
