using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        /// <summary>
        /// Lays out split-row paragraphs independently so a markerless list level's text indent
        /// applies to every wrapped line without shifting adjacent non-list paragraphs.
        /// </summary>
        private static (IReadOnlyList<OfficeRichTextLine> Lines, IReadOnlyList<double> LineIndents) LayoutMarkerlessSplitTableCellParagraphs(
            IReadOnlyList<List<WordParagraph>> paragraphRuns,
            A.ColorScheme? colorScheme,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)>? listMarkers,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            double contentWidth) {
            var lines = new List<OfficeRichTextLine>();
            var indents = new List<double>();
            foreach (List<WordParagraph> runs in paragraphRuns) {
                if (runs.Count == 0) continue;
                List<OfficeRichTextRun> richRuns = CreateSplitTableCellRichRuns(
                    new List<List<WordParagraph>> { runs }, colorScheme, listMarkers, context, diagnostics);
                if (richRuns.Count == 0) continue;

                WordImageListMarker? listMarker = CreateTableCellListMarker(runs, listMarkers);
                double indent = listMarker is { Marker.Length: 0 } marker
                    ? Math.Min(Math.Max(0D, marker.LeftIndentPoints), Math.Max(0D, contentWidth - 1D))
                    : 0D;
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
                    paragraphIndent: null,
                    cancellationToken: context.CancellationToken);
                lines.AddRange(layout.Lines);
                indents.AddRange(Enumerable.Repeat(indent, layout.Lines.Count));
            }

            return (lines, indents);
        }
    }
}
