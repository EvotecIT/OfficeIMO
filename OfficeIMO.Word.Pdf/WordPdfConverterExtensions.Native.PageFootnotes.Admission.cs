using System.Collections.Generic;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    // The PDF owner currently qualifies one single-column body without canvas or
    // non-flow table peers. Select the existing Word fallback before attaching any
    // definitions, including when an unsupported peer is in a different section.
    private static string? GetNativePageFootnoteBodyProfileGap(IReadOnlyList<WordSection> sections,
        WordPdfExportOptions? options) {
        foreach (WordSection section in sections) {
            if ((section.ColumnCount ?? 1) > 1)
                return "The document body contains a multi-column section, outside the current page-note body profile.";
            string? gap = GetNativePageFootnoteElementProfileGap(section.Elements, options, 0, 0);
            if (gap != null) return gap;
        }
        return null;
    }

    private static string? GetNativePageFootnoteElementProfileGap(IEnumerable<WordElement> source,
        WordPdfExportOptions? options, int structuredDepth, int tableDepth) {
        foreach (WordElement element in CollapseNativeParagraphElements(source)) {
            options?.CancellationToken.ThrowIfCancellationRequested();
            if (element is WordStructuredDocumentTag structured) {
                EnsureNativeStructuredDocumentTagDepth(structuredDepth);
                if (TryGetNativeStructuredBlockPropertyValue(structured.Document, structured.SdtBlock, out _)) continue;
                string? gap = GetNativePageFootnoteElementProfileGap(GetNativeStructuredBlockElements(
                    structured.Document, structured.SdtBlock), options, structuredDepth + 1, tableDepth);
                if (gap != null) return gap;
            } else if (element is WordCoverPage cover) {
                EnsureNativeStructuredDocumentTagDepth(structuredDepth);
                if (cover.SdtBlock?.SdtContentBlock is { } content && HasNativeVmlCoverDrawing(content))
                    return "The document body contains a cover-page canvas, outside the current page-note body profile.";
                string? gap = GetNativePageFootnoteElementProfileGap(GetNativeStructuredBlockElements(
                    cover.Document, cover.SdtBlock), options, structuredDepth + 1, tableDepth);
                if (gap != null) return gap;
            } else if (element is WordTable table) {
                EnsureNativeTableDepth(tableDepth);
                bool configured = options?.PdfOptions?.HasExplicitDefaultTableStyle == true;
                PdfCore.PdfTableStyle style = ResolveNativeWordTableStyle(table, configured) ?? CreateNativeDefaultTableStyle(options);
                if (table._tableProperties?.TablePositionProperties != null || style.Position != null || !style.ConsumesVerticalFlow)
                    return "The document body contains a floating or non-flow table, outside the current page-note body profile.";
                foreach (WordTableRow row in table.Rows)
                    foreach (WordTableCell cell in row.Cells) {
                        string? gap = GetNativePageFootnoteElementProfileGap(EnumerateNativeTableCellElements(cell),
                            options, structuredDepth, tableDepth + 1);
                        if (gap != null) return gap;
                    }
            } else if (element is WordParagraph paragraph && !paragraph._paragraph.Ancestors<W.TableCell>().Any()) {
                List<WordParagraph> runs = GetNativeRuns(paragraph, includeInlineContentControls: true)
                    .Where(run => !IsNativeHiddenTextRun(run, paragraph)).ToList();
                if (GetNativeVisibleTextBoxes(paragraph, runs).Count > 0 || HasNativeParagraphShapeGroups(runs) ||
                    runs.SelectMany(run => run.EnumerateEffectiveRunContent()).OfType<W.Drawing>().Any(drawing => drawing.Anchor != null))
                    return "The document body contains anchored graphics, shape groups or text boxes, outside the current page-note body profile.";
            }
        }
        return null;
    }

    private static bool HasNativeStableFootnoteContinuationWidth(IReadOnlyList<WordSection> sections,
        int sourceIndex, WordPdfExportOptions? options) {
        double width = GetNativeFootnoteSectionWidth(sections[sourceIndex], options);
        for (int index = sourceIndex + 1; index < sections.Count; index++)
            if (!NativeDoubleEquals(width, GetNativeFootnoteSectionWidth(sections[index], options))) return false;
        return true;
    }

    private static double GetNativeFootnoteSectionWidth(WordSection section, WordPdfExportOptions? options) {
        PdfCore.PdfPageMargins margins = GetNativeMargins(section, options, (0D, 0D));
        return Math.Max(1D, GetNativePageSize(section, options).Width.ToPoints() - margins.Left.ToPoints() - margins.Right.ToPoints());
    }
}
