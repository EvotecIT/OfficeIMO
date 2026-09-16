using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        /// <summary>
        /// Lays out list paragraphs in split rows independently so marker indentation, picture-marker
        /// diagnostics, and non-text content remain associated with the fragment that consumes them.
        /// </summary>
        private static (IReadOnlyList<OfficeRichTextLine> Lines, IReadOnlyList<double> LineIndents,
            IReadOnlyList<SplitTableCellContentEntry> ContentOrder) LayoutListSplitTableCellParagraphs(
            WordTableCell cell,
            A.ColorScheme? colorScheme,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            double contentWidth) {
            var lines = new List<OfficeRichTextLine>();
            var indents = new List<double>();
            var contentOrder = new List<SplitTableCellContentEntry>();
            int imageIndex = 0;
            int nestedTableIndex = 0;
            foreach (var child in cell._tableCell.ChildElements) {
                context.ThrowIfCancellationRequested();
                if (child is Table) {
                    contentOrder.Add(SplitTableCellContentEntry.CreateNestedTable(nestedTableIndex++));
                    continue;
                }

                if (child is not Paragraph paragraph) continue;
                WordImageListMarker? listMarker = listMarkers == null
                    ? null
                    : CreateListMarker(cell.Document, paragraph, listMarkers);
                double textOffset = listMarker.HasValue
                    ? Math.Min(Math.Max(0D, listMarker.Value.LeftIndentPoints), Math.Max(0D, contentWidth - 1D))
                    : 0D;
                double markerOffset = listMarker is { Marker.Length: > 0 } visibleMarker
                    ? Math.Max(0D, textOffset - Math.Max(0D, visibleMarker.HangingIndentPoints))
                    : textOffset;
                double indent = markerOffset;
                OfficeTextParagraphIndent paragraphIndent = listMarker is { Marker.Length: > 0 }
                    ? OfficeTextParagraphIndent.Hanging(Math.Max(0D, textOffset - markerOffset))
                    : OfficeTextParagraphIndent.Empty;
                bool markerPending = listMarker is { Marker.Length: > 0 };
                var segmentRuns = new List<WordParagraph>();

                void FlushTextSegment() {
                    List<OfficeRichTextRun> richRuns = CreateRichTextRuns(segmentRuns, colorScheme, context, diagnostics);
                    segmentRuns.Clear();
                    int? pictureBulletId = null;
                    if (markerPending && listMarker is { Marker.Length: > 0 } visible) {
                        richRuns.Insert(0, CreateListMarkerRichTextRun(visible, contentWidth));
                        pictureBulletId = visible.PictureBulletId;
                        markerPending = false;
                    }
                    if (richRuns.Count == 0) return;

                    double maxFontSize = richRuns.Max(run => run.FontSize);
                    double lineHeight = Math.Max(maxFontSize * 1.25D, 12D);
                    OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
                        richRuns,
                        Math.Max(1D, contentWidth - indent),
                        double.MaxValue,
                        Math.Max(1D, lineHeight / Math.Max(1D, maxFontSize)),
                        CreateRichTextMeasure(context.CancellationToken),
                        wrap: true,
                        shrinkToFit: false,
                        minimumFontSize: Math.Min(6D, maxFontSize),
                        overflowBehavior: OfficeTextOverflowBehavior.Clip,
                        paragraphIndent: paragraphIndent,
                        cancellationToken: context.CancellationToken);
                    contentOrder.Add(SplitTableCellContentEntry.CreateText(lines.Count, layout.Lines.Count, pictureBulletId));
                    lines.AddRange(layout.Lines);
                    indents.AddRange(layout.Lines.Select(line => indent + line.OffsetX));
                }

                foreach (WordParagraph run in WordSection.ConvertParagraphToWordParagraphs(cell.Document, paragraph, splitPaginationMarkers: true,
                    cancellationToken: context.CancellationToken)) {
                    if (run.IsPageBreak || run.IsColumnBreak) continue;
                    if (run.Image != null) {
                        FlushTextSegment();
                        contentOrder.Add(SplitTableCellContentEntry.CreateImage(imageIndex++));
                    } else {
                        segmentRuns.Add(run);
                    }
                }
                FlushTextSegment();
            }

            return (lines, indents, contentOrder);
        }
    }
}
