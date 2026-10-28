using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    // Keep paragraph boundaries in the PDF owner's existing paragraph model;
    // flattening them to newlines discards leading, blank marks and spacing.
    private static IReadOnlyList<PdfCore.PdfTableCellParagraph> CreatePdfTableCellParagraphs(
        PptCore.PowerPointTableCell cell, string? fallbackFontFamily, double fallbackFontSize,
        double fallbackLineHeight, int slideNumber, PowerPointPdfExportOptions options,
        IReadOnlyList<A.TextCharacterPropertiesType> tableDefaults) {
        var paragraphs = new List<PdfCore.PdfTableCellParagraph>();
        if (cell.Cell.TextBody == null) return paragraphs;
        var numbering = new PptCore.PowerPointParagraphNumberingState();
        bool reportedListLayout = false;
        foreach (A.Paragraph paragraph in cell.Cell.TextBody.Elements<A.Paragraph>()) {
            options.CancellationToken.ThrowIfCancellationRequested();
            var wrapped = new PptCore.PowerPointParagraph(paragraph, cell.SlidePart, cell.OwnerPart);
            IReadOnlyList<A.TextCharacterPropertiesType> sources = ResolveTableParagraphTextPropertySources(cell, paragraph, tableDefaults);
            double fontSize = ReadRunFontSize(sources) ?? fallbackFontSize;
            var runs = new List<PdfCore.PdfTextRun>();
            if (numbering.CreateMarker(wrapped) is string marker) {
                runs.Add(CreatePdfTableCellListMarker(cell, wrapped, marker,
                    fallbackFontFamily, fontSize, slideNumber, options, tableDefaults));
                if (!reportedListLayout) {
                    AddWarning(options, slideNumber, "table-list-layout-approximation",
                        "PowerPoint table list markers and counters are retained as inline text. Marker positioning is approximated.");
                    reportedListLayout = true;
                }
            }
            AppendPdfTableCellParagraphRuns(runs, paragraph, cell, fallbackFontFamily, fallbackFontSize, slideNumber, options, tableDefaults);
            if (runs.Count == 0 || runs.All(run => string.IsNullOrEmpty(run.Text))) {
                string? family = ReadRunFontName(cell, sources) ?? fallbackFontFamily;
                runs.Clear();
                runs.Add(new PdfCore.PdfTextRun(string.Empty, fontSize: fontSize, font: MapFont(family), fontFamily: family));
            }
            PdfCore.PdfLineSpacing lineSpacing = wrapped.LineSpacing?.ToPoints() is double points
                ? PdfCore.PdfLineSpacing.Exactly(Math.Max(1D, points))
                : PdfCore.PdfLineSpacing.Multiple(Math.Max(.01D, wrapped.LineSpacingMultiplier ?? fallbackLineHeight));
            paragraphs.Add(new PdfCore.PdfTableCellParagraph(runs,
                spacingBefore: Math.Max(0D, wrapped.SpaceBefore?.ToPoints() ?? 0D),
                spacingAfter: Math.Max(0D, wrapped.SpaceAfter?.ToPoints() ?? 0D),
                align: MapAlign((wrapped.Alignment ?? cell.HorizontalAlignment).ToOpenXml()),
                leftIndent: wrapped.LeftMargin?.ToPoints() ?? 0D,
                firstLineIndent: wrapped.Indent?.ToPoints() ?? 0D,
                fontSize: fontSize, lineSpacing: lineSpacing));
        }
        return paragraphs;
    }

    private static IReadOnlyList<A.TextCharacterPropertiesType> ResolveTableParagraphTextPropertySources(
        PptCore.PowerPointTableCell cell, A.Paragraph paragraph, IReadOnlyList<A.TextCharacterPropertiesType> tableDefaults) {
        bool blank = !paragraph.Descendants<A.Text>().Any(text => !string.IsNullOrEmpty(text.Text));
        A.TextCharacterPropertiesType? first = paragraph.ChildElements.Select(child => child switch {
            A.Run run => (A.TextCharacterPropertiesType?)run.RunProperties,
            A.Field field => field.RunProperties,
            _ => null
        }).FirstOrDefault(properties => properties != null);
        A.TextCharacterPropertiesType? mark = blank ? paragraph.GetFirstChild<A.EndParagraphRunProperties>() : null;
        var sources = ResolveTableTextPropertySources(cell, paragraph, mark ?? first, tableDefaults).ToList();
        if (mark != null && first != null) sources.Insert(1, first);
        return sources;
    }
}
