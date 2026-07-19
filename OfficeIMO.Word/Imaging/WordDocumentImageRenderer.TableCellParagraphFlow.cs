using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private static void AddTableCellParagraphFlow(
            WordTableCell cell,
            OfficeDrawing drawing,
            double contentLeft,
            double contentTop,
            double contentWidth,
            double contentHeight,
            double contentBottom,
            IReadOnlyList<IReadOnlyList<WordParagraph>> paragraphRuns,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers,
            A.ColorScheme? colorScheme,
            List<OfficeImageExportDiagnostic> diagnostics,
            WordImageFlowContext? parentContext = null) {
            CancellationToken cancellationToken = parentContext?.CancellationToken ?? default;
            cancellationToken.ThrowIfCancellationRequested();
            double flowHeight = EstimateTableCellParagraphFlowHeight(
                paragraphRuns,
                contentWidth,
                listMarkers,
                colorScheme,
                cancellationToken);
            double flowTop = contentTop + ResolveTableCellVerticalOffset(cell.VerticalAlignment, contentHeight, flowHeight);
            WordImageFlowContext context = CreateFlowContext(
                drawing,
                contentLeft,
                flowTop,
                contentWidth,
                contentBottom,
                "unsupported-word-table-cell-text-overflow",
                "Skipped Word table cell text because it does not fit within the cell content area.",
                resolveDynamicPageFields: parentContext?.ResolveDynamicPageFields ?? false,
                totalPageCount: parentContext?.TotalPageCount ?? 1,
                sectionNumber: parentContext?.SectionNumber ?? 1,
                sectionPageCount: parentContext?.SectionPageCount ?? 1,
                pageNumberValue: parentContext?.PageNumberValue ?? 0,
                pageNumberText: parentContext?.PageNumberText,
                cancellationToken: cancellationToken);

            AddTableCellParagraphRuns(paragraphRuns, context, diagnostics, listMarkers, colorScheme);
        }

        private static double EstimateTableCellTextHeight(
            WordTableCell cell,
            IReadOnlyList<IReadOnlyList<WordParagraph>> paragraphRuns,
            double fontSize,
            double width,
            double lineHeight,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            if (paragraphRuns.Count > 1) {
                return EstimateTableCellParagraphFlowHeight(
                    paragraphRuns,
                    width,
                    listMarkers,
                    GetDocumentColorScheme(cell.Document),
                    cancellationToken);
            }

            string text = GetCellText(
                cell,
                cancellationToken: cancellationToken);
            return string.IsNullOrWhiteSpace(text)
                ? lineHeight
                : EstimateTextHeight(
                    text,
                    fontSize,
                    width,
                    lineHeight,
                    cancellationToken);
        }

        private static double EstimateTableCellParagraphFlowHeight(
            IReadOnlyList<IReadOnlyList<WordParagraph>> paragraphRuns,
            double contentWidth,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers,
            A.ColorScheme? colorScheme,
            CancellationToken cancellationToken = default) {
            var measurementDrawing = new OfficeDrawing(Math.Max(1D, contentWidth), double.MaxValue);
            WordImageFlowContext measurementContext = CreateFlowContext(
                measurementDrawing,
                0D,
                0D,
                contentWidth,
                double.MaxValue,
                "unsupported-word-table-cell-measurement-overflow",
                "Skipped Word table cell text while measuring table cell content height.",
                cancellationToken: cancellationToken);
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            AddTableCellParagraphRuns(paragraphRuns, measurementContext, diagnostics, listMarkers, colorScheme);
            return Math.Max(0D, measurementContext.Y);
        }

        private static void AddTableCellParagraphRuns(
            IReadOnlyList<IReadOnlyList<WordParagraph>> paragraphRuns,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers,
            A.ColorScheme? colorScheme) {
            for (int i = 0; i < paragraphRuns.Count; i++) {
                context.ThrowIfCancellationRequested();
                IReadOnlyList<WordParagraph> runs = paragraphRuns[i];
                WordImageListMarker? listMarker = CreateTableCellListMarker(runs, listMarkers);
                bool added = runs.Count == 1 && !HasRunHighlight(runs[0])
                    ? AddTextRun(runs[0], context, diagnostics, listMarker, colorScheme)
                    : AddRichTextRuns(runs, context, diagnostics, listMarker, colorScheme);
                if (!added || context.StoppedForPagination) {
                    break;
                }
            }
        }

        private static WordImageListMarker? CreateTableCellListMarker(
            IReadOnlyList<WordParagraph> runs,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers) {
            if (runs.Count == 0 || listMarkers == null) {
                return null;
            }

            WordParagraph firstRun = runs[0];
            return CreateListMarker(firstRun._document, firstRun._paragraph, listMarkers);
        }

        private static List<List<WordParagraph>> CreateTableCellParagraphRuns(
            WordTableCell cell,
            CancellationToken cancellationToken = default) {
            var paragraphs = new List<List<WordParagraph>>();
            foreach (Paragraph paragraph in cell._tableCell.ChildElements.OfType<Paragraph>()) {
                cancellationToken.ThrowIfCancellationRequested();
                List<WordParagraph> paragraphRuns = WordSection.ConvertParagraphToWordParagraphs(
                        cell.Document,
                        paragraph,
                        splitPaginationMarkers: true,
                        cancellationToken)
                    .Where(run => !run.IsPageBreak && !run.IsColumnBreak)
                    .Where(run => !string.IsNullOrEmpty(run.Text))
                    .ToList();
                if (paragraphRuns.Count > 0) {
                    paragraphs.Add(paragraphRuns);
                }
            }

            return paragraphs;
        }

        private static double ResolveTableCellVerticalOffset(TableVerticalAlignmentValues? alignment, double contentHeight, double flowHeight) {
            double extraHeight = Math.Max(0D, contentHeight - flowHeight);
            if (alignment == TableVerticalAlignmentValues.Center) {
                return extraHeight / 2D;
            }

            if (alignment == TableVerticalAlignmentValues.Bottom) {
                return extraHeight;
            }

            return 0D;
        }
    }
}
