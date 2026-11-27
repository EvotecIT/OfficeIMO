using System.Globalization;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static bool ShouldExportSections(WordDocument document) {
        return document.Sections.Count > 1 ||
            document.Sections.Any(section =>
                section.ColumnCount.HasValue ||
                section.ColumnsSpace.HasValue ||
                section.HasColumnSeparator ||
                section.Node().WChild("cols").WElements("col").Any() ||
                section.Node().WChild("vAlign").Exists ||
                section.Node().WChild("lnNumType").Exists ||
                section.Node().WChild("type").Exists);
    }

    private static void CopySections(WordDocument document, RtfDocument rtf, Dictionary<string, int> revisionAuthorIndexes) {
        for (int index = 0; index < document.Sections.Count; index++) {
            WordSection wordSection = document.Sections[index];
            RtfSection section = rtf.AddSection(ToRtfSectionBreakKind(document, index));
            CopyPageSetup(wordSection, section, rtf);
            CopySectionHeaderFooters(wordSection, section, rtf, revisionAuthorIndexes);
            CopyWordElements(wordSection.Elements, section, rtf, revisionAuthorIndexes);
        }
    }

    private static RtfSectionBreakKind ToRtfSectionBreakKind(WordDocument document, int sectionIndex) {
        WordSection section = document.Sections[sectionIndex];
        string? sectionMark = section.Node().WChild("type").WVal();
        if (sectionMark == "continuous") return RtfSectionBreakKind.Continuous;
        if (sectionMark == "nextColumn") return RtfSectionBreakKind.Column;
        if (sectionMark == "evenPage") return RtfSectionBreakKind.EvenPage;
        if (sectionMark == "oddPage") return RtfSectionBreakKind.OddPage;
        return RtfSectionBreakKind.NextPage;
    }

    private static void CopyPageSetup(WordSection source, RtfSection destination, RtfDocument document) {
        int? width = ToInt32(checked((uint?)(source.PageSetup.Size.Width?.ToTwips())));
        int? height = ToInt32(checked((uint?)(source.PageSetup.Size.Height?.ToTwips())));
        if (width.HasValue && height.HasValue) {
            destination.PageSetup.SetPaperSize(width.Value, height.Value);
        } else {
            destination.PageSetup.PaperWidthTwips = width;
            destination.PageSetup.PaperHeightTwips = height;
        }

        destination.PageSetup.SetMargins(
            ToInt32(checked((uint)(source.PageSetup.Margins.Left.GetValueOrDefault().ToTwips()))),
            ToInt32(checked((uint)(source.PageSetup.Margins.Right.GetValueOrDefault().ToTwips()))),
            checked((int?)(source.PageSetup.Margins.Top?.ToTwips())),
            checked((int?)(source.PageSetup.Margins.Bottom?.ToTwips())));
        destination.PageSetup.SetGutter(ToInt32(checked((uint)(source.PageSetup.Margins.Gutter.GetValueOrDefault().ToTwips()))), source.RtlGutter);
        destination.PageSetup.SetHeaderFooterDistance(
            ToInt32(checked((uint)(source.PageSetup.Margins.HeaderDistance.GetValueOrDefault().ToTwips()))),
            ToInt32(checked((uint)(source.PageSetup.Margins.FooterDistance.GetValueOrDefault().ToTwips()))));
        destination.PageSetup.SetLandscape(source.PageSetup.Orientation == OfficePageOrientation.Landscape);
        destination.PageSetup.SetDifferentFirstPageHeaderFooter(source.PageSetup.DifferentFirstPage);
        OfficeXmlNode sectionProperties = source.Node();
        CopyPageNumbering(sectionProperties.WChild("pgNumType"), destination.PageSetup);
        CopyPageBorders(sectionProperties.WChild("pgBorders"), destination.PageSetup.PageBorders, document);
        CopyNoteSettings(
            sectionProperties.WChild("footnotePr"),
            sectionProperties.WChild("endnotePr"),
            destination.NoteSettings);
        CopyLineNumbering(sectionProperties.WChild("lnNumType"), destination.LineNumbering);
        destination.VerticalAlignment = ToRtfSectionVerticalAlignment(sectionProperties.WChild("vAlign").WVal());
        destination.ColumnCount = source.ColumnCount;
        destination.ColumnSpaceTwips = checked((int?)(source.ColumnsSpace?.ToTwips()));
        destination.ColumnSeparator = source.HasColumnSeparator;
        CopySectionColumns(sectionProperties.WChild("cols"), destination);
    }

    private static void ApplySections(RtfDocument rtfDocument, WordDocument document, RtfWordNumberingPlan numbering) {
        var wordSections = new WordSection[rtfDocument.Sections.Count];
        wordSections[0] = document.Sections[0];

        for (int index = 0; index < rtfDocument.Sections.Count; index++) {
            RtfSection rtfSection = rtfDocument.Sections[index];
            WordSection wordSection = index == 0
                ? wordSections[0]
                : document.AddSection(ToWordSectionMark(rtfSection.BreakKind));
            wordSections[index] = wordSection;

            if (index == 0) {
                ApplySectionBreakKind(rtfSection.BreakKind, wordSection);
            }

            foreach (IRtfBlock block in rtfSection.Blocks) {
                if (block is RtfParagraph paragraph) {
                    AppendParagraph(wordSection, paragraph, rtfDocument, numbering);
                } else if (block is RtfTable table) {
                    AppendTable(wordSection, table, rtfDocument, numbering);
                } else if (block is RtfImage image) {
                    AppendImage(wordSection, image);
                }
            }
        }

        for (int index = 0; index < rtfDocument.Sections.Count; index++) {
            ApplyPageSetup(rtfDocument.Sections[index], wordSections[index], rtfDocument);
        }
        ApplySectionHeaderFooters(rtfDocument, document, wordSections, numbering);
    }

    private static void ApplyPageSetup(RtfSection source, WordSection destination, RtfDocument rtfDocument) {
        RtfPageSetup pageSetup = rtfDocument.GetPageSetup(source);
        if (pageSetup.DirectLandscape.HasValue)
            destination.PageSetup.Orientation = pageSetup.Landscape ? OfficePageOrientation.Landscape : OfficePageOrientation.Portrait;

        if (pageSetup.PaperWidthTwips.HasValue) {
            destination.PageSetup.Size.Width= ((ToUInt32Value(pageSetup.PaperWidthTwips.Value)) is { } wordLength6282 ? global::OfficeIMO.OfficeLength.Points((wordLength6282) / 20D) : (global::OfficeIMO.OfficeLength?)null);
        }

        if (pageSetup.PaperHeightTwips.HasValue) {
            destination.PageSetup.Size.Height= ((ToUInt32Value(pageSetup.PaperHeightTwips.Value)) is { } wordLength6438 ? global::OfficeIMO.OfficeLength.Points((wordLength6438) / 20D) : (global::OfficeIMO.OfficeLength?)null);
        }

        if (pageSetup.MarginLeftTwips.HasValue) {
            destination.PageSetup.Margins.Left = global::OfficeIMO.OfficeLength.Points((ToUInt32Value(pageSetup.MarginLeftTwips.Value)) / 20D);
        }

        if (pageSetup.MarginRightTwips.HasValue) {
            destination.PageSetup.Margins.Right = global::OfficeIMO.OfficeLength.Points((ToUInt32Value(pageSetup.MarginRightTwips.Value)) / 20D);
        }

        if (pageSetup.MarginTopTwips.HasValue) {
            destination.PageSetup.Margins.Top= ((pageSetup.MarginTopTwips.Value) is { } wordLength6884 ? global::OfficeIMO.OfficeLength.Points((wordLength6884) / 20D) : (global::OfficeIMO.OfficeLength?)null);
        }

        if (pageSetup.MarginBottomTwips.HasValue) {
            destination.PageSetup.Margins.Bottom= ((pageSetup.MarginBottomTwips.Value) is { } wordLength7020 ? global::OfficeIMO.OfficeLength.Points((wordLength7020) / 20D) : (global::OfficeIMO.OfficeLength?)null);
        }

        if (pageSetup.GutterWidthTwips.HasValue) {
            destination.PageSetup.Margins.Gutter = global::OfficeIMO.OfficeLength.Points((ToUInt32Value(pageSetup.GutterWidthTwips.Value)) / 20D);
        }

        if (pageSetup.HeaderDistanceTwips.HasValue) {
            destination.PageSetup.Margins.HeaderDistance = global::OfficeIMO.OfficeLength.Points((ToUInt32Value(pageSetup.HeaderDistanceTwips.Value)) / 20D);
        }

        if (pageSetup.FooterDistanceTwips.HasValue) {
            destination.PageSetup.Margins.FooterDistance = global::OfficeIMO.OfficeLength.Points((ToUInt32Value(pageSetup.FooterDistanceTwips.Value)) / 20D);
        }

        ApplyFirstPageHeaderFooterSelection(destination, pageSetup.DifferentFirstPageHeaderFooter);
        destination.RtlGutter = pageSetup.RtlGutter;
        ApplyPageNumbering(pageSetup, destination);
        ApplyPageBorders(pageSetup.PageBorders, destination, rtfDocument);
        ApplyNoteSettings(RtfNoteProjection.Resolve(rtfDocument, source).Settings, destination);
        ApplyLineNumbering(source.LineNumbering, destination);
        ApplySectionVerticalAlignment(source.VerticalAlignment, destination);
        destination.ColumnCount = source.ColumnCount;
        destination.ColumnsSpace = source.ColumnSpaceTwips.HasValue
            ? OfficeLength.Points(source.ColumnSpaceTwips.Value / 20D) : (OfficeLength?)null;
        destination.HasColumnSeparator = source.ColumnSeparator;
        ApplySectionColumns(source, destination);
    }

    private static void CopySectionColumns(OfficeXmlNode source, RtfSection destination) {
        if (source.IsNone) {
            return;
        }

        foreach (OfficeXmlNode column in source.WElements("col")) {
            destination.AddColumn(ToInt32(column.WAttr("w")), ToInt32(column.WAttr("space")));
        }

        if (!destination.ColumnCount.HasValue && destination.Columns.Count > 0) {
            destination.ColumnCount = destination.Columns.Count;
        }
    }

    private static void ApplySectionColumns(RtfSection source, WordSection destination) {
        if (source.Columns.Count == 0) {
            return;
        }

        OfficeXmlNode sectionProperties = destination.Node();
        OfficeXmlNode columns = sectionProperties.WChild("cols");
        if (columns.IsNone) {
            columns = sectionProperties.Append(sectionProperties.CreateElement(WordXml.W, "cols"));
        }

        columns.SetAttribute(WordXml.W, "equalWidth", FormatOnOff(false));
        columns.SetAttribute(WordXml.W, "num", ((short)(source.ColumnCount ?? source.Columns.Count)).ToString(CultureInfo.InvariantCulture));
        columns.RemoveChildren(WordXml.W, "col");
        foreach (RtfSectionColumn sourceColumn in source.Columns) {
            OfficeXmlNode column = columns.CreateElement(WordXml.W, "col");
            if (sourceColumn.WidthTwips.HasValue) {
                column.SetAttribute(WordXml.W, "w", sourceColumn.WidthTwips.Value.ToString(CultureInfo.InvariantCulture));
            }

            if (sourceColumn.SpaceAfterTwips.HasValue) {
                column.SetAttribute(WordXml.W, "space", sourceColumn.SpaceAfterTwips.Value.ToString(CultureInfo.InvariantCulture));
            }

            columns.Append(column);
        }
    }

    private static void ApplySectionBreakKind(RtfSectionBreakKind kind, WordSection section) {
        section.BreakType = ToWordSectionMark(kind);
    }

    private static WordSectionBreakType ToWordSectionMark(RtfSectionBreakKind kind) {
        switch (kind) {
            case RtfSectionBreakKind.Continuous:
                return WordSectionBreakType.Continuous;
            case RtfSectionBreakKind.Column:
                return WordSectionBreakType.NextColumn;
            case RtfSectionBreakKind.EvenPage:
                return WordSectionBreakType.EvenPage;
            case RtfSectionBreakKind.OddPage:
                return WordSectionBreakType.OddPage;
            default:
                return WordSectionBreakType.NextPage;
        }
    }

    private static void AppendParagraph(WordSection section, RtfParagraph paragraph, RtfDocument rtfDocument, RtfWordNumberingPlan numbering) {
        WordParagraph wordParagraph = section.AddParagraph();
        ApplyParagraphFormatting(wordParagraph, paragraph, rtfDocument, numbering);
        AppendRuns(wordParagraph, paragraph, rtfDocument, numbering);
    }

}
