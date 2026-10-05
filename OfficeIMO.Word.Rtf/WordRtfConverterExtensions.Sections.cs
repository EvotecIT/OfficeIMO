using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static bool ShouldExportSections(WordDocument document) {
        return document.Sections.Count > 1 ||
            document.Sections.Any(section =>
                section.ColumnCount.HasValue ||
                section.ColumnsSpace.HasValue ||
                section.HasColumnSeparator ||
                section._sectionProperties.GetFirstChild<Columns>()?.Elements<Column>().Any() == true ||
                section._sectionProperties.GetFirstChild<VerticalTextAlignmentOnPage>() != null ||
                section._sectionProperties.GetFirstChild<LineNumberType>() != null ||
                section._sectionProperties.GetFirstChild<SectionType>() != null);
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
        SectionMarkValues? sectionMark = section._sectionProperties.GetFirstChild<SectionType>()?.Val?.Value;
        if (sectionMark == SectionMarkValues.Continuous) return RtfSectionBreakKind.Continuous;
        if (sectionMark == SectionMarkValues.NextColumn) return RtfSectionBreakKind.Column;
        if (sectionMark == SectionMarkValues.EvenPage) return RtfSectionBreakKind.EvenPage;
        if (sectionMark == SectionMarkValues.OddPage) return RtfSectionBreakKind.OddPage;
        return RtfSectionBreakKind.NextPage;
    }

    private static void CopyPageSetup(WordSection source, RtfSection destination, RtfDocument document) {
        int? width = ToInt32(source.PageSettings.Width);
        int? height = ToInt32(source.PageSettings.Height);
        if (width.HasValue && height.HasValue) {
            destination.PageSetup.SetPaperSize(width.Value, height.Value);
        } else {
            destination.PageSetup.PaperWidthTwips = width;
            destination.PageSetup.PaperHeightTwips = height;
        }

        destination.PageSetup.SetMargins(
            ToInt32(source.Margins.Left),
            ToInt32(source.Margins.Right),
            source.Margins.Top,
            source.Margins.Bottom);
        destination.PageSetup.SetGutter(ToInt32(source.Margins.Gutter), source.RtlGutter);
        destination.PageSetup.SetHeaderFooterDistance(
            ToInt32(source.Margins.HeaderDistance),
            ToInt32(source.Margins.FooterDistance));
        destination.PageSetup.SetLandscape(source.PageOrientation == OfficePageOrientation.Landscape);
        destination.PageSetup.SetDifferentFirstPageHeaderFooter(source.DifferentFirstPage);
        CopyPageNumbering(source._sectionProperties.GetFirstChild<PageNumberType>(), destination.PageSetup);
        CopyPageBorders(source._sectionProperties.GetFirstChild<PageBorders>(), destination.PageSetup.PageBorders, document);
        CopyNoteSettings(
            source._sectionProperties.GetFirstChild<FootnoteProperties>(),
            source._sectionProperties.GetFirstChild<EndnoteProperties>(),
            destination.NoteSettings);
        CopyLineNumbering(source._sectionProperties.GetFirstChild<LineNumberType>(), destination.LineNumbering);
        destination.VerticalAlignment = ToRtfSectionVerticalAlignment(source._sectionProperties.GetFirstChild<VerticalTextAlignmentOnPage>()?.Val?.Value);
        destination.ColumnCount = source.ColumnCount;
        destination.ColumnSpaceTwips = source.ColumnsSpace;
        destination.ColumnSeparator = source.HasColumnSeparator;
        CopySectionColumns(source._sectionProperties.GetFirstChild<Columns>(), destination);
    }

    private static void ApplySections(RtfDocument rtfDocument, WordDocument document) {
        var wordSections = new WordSection[rtfDocument.Sections.Count];
        wordSections[0] = document.Sections[0];

        for (int index = 0; index < rtfDocument.Sections.Count; index++) {
            RtfSection rtfSection = rtfDocument.Sections[index];
            WordSection wordSection = index == 0
                ? wordSections[0]
                : document.AddSection(ToWordSectionMark(rtfSection.BreakKind).ToOfficeEnum());
            wordSections[index] = wordSection;

            if (index == 0) {
                ApplySectionBreakKind(rtfSection.BreakKind, wordSection);
            }

            foreach (IRtfBlock block in rtfSection.Blocks) {
                if (block is RtfParagraph paragraph) {
                    AppendParagraph(wordSection, paragraph, rtfDocument);
                } else if (block is RtfTable table) {
                    AppendTable(wordSection, table, rtfDocument);
                } else if (block is RtfImage image) {
                    AppendImage(wordSection, image);
                }
            }
        }

        for (int index = 0; index < rtfDocument.Sections.Count; index++) {
            ApplyPageSetup(rtfDocument.Sections[index], wordSections[index], rtfDocument);
        }
        ApplySectionHeaderFooters(rtfDocument, document, wordSections);
    }

    private static void ApplyPageSetup(RtfSection source, WordSection destination, RtfDocument rtfDocument) {
        RtfPageSetup pageSetup = rtfDocument.GetPageSetup(source);
        if (pageSetup.DirectLandscape.HasValue)
            destination.PageOrientation = pageSetup.Landscape ? OfficePageOrientation.Landscape : OfficePageOrientation.Portrait;

        if (pageSetup.PaperWidthTwips.HasValue) {
            destination.PageSettings.Width = ToUInt32Value(pageSetup.PaperWidthTwips.Value);
        }

        if (pageSetup.PaperHeightTwips.HasValue) {
            destination.PageSettings.Height = ToUInt32Value(pageSetup.PaperHeightTwips.Value);
        }

        if (pageSetup.MarginLeftTwips.HasValue) {
            destination.Margins.Left = ToUInt32Value(pageSetup.MarginLeftTwips.Value);
        }

        if (pageSetup.MarginRightTwips.HasValue) {
            destination.Margins.Right = ToUInt32Value(pageSetup.MarginRightTwips.Value);
        }

        if (pageSetup.MarginTopTwips.HasValue) {
            destination.Margins.Top = pageSetup.MarginTopTwips.Value;
        }

        if (pageSetup.MarginBottomTwips.HasValue) {
            destination.Margins.Bottom = pageSetup.MarginBottomTwips.Value;
        }

        if (pageSetup.GutterWidthTwips.HasValue) {
            destination.Margins.Gutter = ToUInt32Value(pageSetup.GutterWidthTwips.Value);
        }

        if (pageSetup.HeaderDistanceTwips.HasValue) {
            destination.Margins.HeaderDistance = ToUInt32Value(pageSetup.HeaderDistanceTwips.Value);
        }

        if (pageSetup.FooterDistanceTwips.HasValue) {
            destination.Margins.FooterDistance = ToUInt32Value(pageSetup.FooterDistanceTwips.Value);
        }

        ApplyFirstPageHeaderFooterSelection(destination, pageSetup.DifferentFirstPageHeaderFooter);
        destination.RtlGutter = pageSetup.RtlGutter;
        ApplyPageNumbering(pageSetup, destination);
        ApplyPageBorders(pageSetup.PageBorders, destination, rtfDocument);
        ApplyNoteSettings(source.NoteSettings, destination);
        ApplyLineNumbering(source.LineNumbering, destination);
        ApplySectionVerticalAlignment(source.VerticalAlignment, destination);
        destination.ColumnCount = source.ColumnCount;
        destination.ColumnsSpace = source.ColumnSpaceTwips;
        destination.HasColumnSeparator = source.ColumnSeparator;
        ApplySectionColumns(source, destination);
    }

    private static void CopySectionColumns(Columns? source, RtfSection destination) {
        if (source == null) {
            return;
        }

        foreach (Column column in source.Elements<Column>()) {
            destination.AddColumn(ToInt32(column.Width), ToInt32(column.Space));
        }

        if (!destination.ColumnCount.HasValue && destination.Columns.Count > 0) {
            destination.ColumnCount = destination.Columns.Count;
        }
    }

    private static void ApplySectionColumns(RtfSection source, WordSection destination) {
        if (source.Columns.Count == 0) {
            return;
        }

        Columns? columns = destination._sectionProperties.GetFirstChild<Columns>();
        if (columns == null) {
            columns = new Columns();
            destination._sectionProperties.Append(columns);
        }

        columns.EqualWidth = false;
        columns.ColumnCount = (Int16Value)(short)(source.ColumnCount ?? source.Columns.Count);
        columns.RemoveAllChildren<Column>();
        foreach (RtfSectionColumn sourceColumn in source.Columns) {
            var column = new Column();
            if (sourceColumn.WidthTwips.HasValue) {
                column.Width = sourceColumn.WidthTwips.Value.ToString(CultureInfo.InvariantCulture);
            }

            if (sourceColumn.SpaceAfterTwips.HasValue) {
                column.Space = sourceColumn.SpaceAfterTwips.Value.ToString(CultureInfo.InvariantCulture);
            }

            columns.Append(column);
        }
    }

    private static void ApplySectionBreakKind(RtfSectionBreakKind kind, WordSection section) {
        section._sectionProperties.RemoveAllChildren<SectionType>();
        section._sectionProperties.Append(new SectionType { Val = ToWordSectionMark(kind) });
    }

    private static SectionMarkValues ToWordSectionMark(RtfSectionBreakKind kind) {
        switch (kind) {
            case RtfSectionBreakKind.Continuous:
                return SectionMarkValues.Continuous;
            case RtfSectionBreakKind.Column:
                return SectionMarkValues.NextColumn;
            case RtfSectionBreakKind.EvenPage:
                return SectionMarkValues.EvenPage;
            case RtfSectionBreakKind.OddPage:
                return SectionMarkValues.OddPage;
            default:
                return SectionMarkValues.NextPage;
        }
    }

    private static void AppendParagraph(WordSection section, RtfParagraph paragraph, RtfDocument rtfDocument) {
        WordParagraph wordParagraph = section.AddParagraph();
        ApplyTabStops(wordParagraph, paragraph);
        ApplyParagraphFormatting(wordParagraph, paragraph, rtfDocument);
        AppendRuns(wordParagraph, paragraph, rtfDocument);
    }

}
