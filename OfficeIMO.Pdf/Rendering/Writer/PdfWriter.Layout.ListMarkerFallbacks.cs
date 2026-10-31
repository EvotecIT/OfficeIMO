namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed class ListMarkerFallbackLayout {
        internal IReadOnlyList<PdfTextRun> Runs = null!;
        internal List<List<RichSeg>> Lines = null!;
        internal List<double> Heights = null!;
        internal double Width;
    }

    private sealed partial class LayoutContext {
        /// <summary>Uses the ordinary run fallback planner for markers before measuring their column.</summary>
        private ListMarkerFallbackLayout? PrepareListMarkerFallback(string marker, PdfStandardFont font,
            double size, double leading, PdfListStyle? style) {
            var run = new PdfTextRun(marker, fontSize: global::OfficeIMO.OfficeLength.Points(size), font: font, fontFamily: style?.MarkerFontFamily)
                .WithInheritedFontStyle(font == ChooseBold(font) || font == ChooseBoldItalic(font),
                    font == ChooseItalic(font) || font == ChooseBoldItalic(font));
            if (style?.MarkerFeatureSettings is { } features) run = run.WithFeatureSettings(features);
            IReadOnlyList<PdfTextRun>? runs = ResolveRunFallbackExpansion(run, font, currentOpts);
            if (runs == null) return null;
            var layout = WrapRichRunsCore(runs, int.MaxValue, size, font, leading, null,
                DefaultParagraphTabStopWidth, currentOpts);
            return new ListMarkerFallbackLayout {
                Runs = runs,
                // The caller supplies the resolved body baseline. Positioned lines
                // retain it instead of applying paragraph ascent resolution again.
                Lines = layout.Lines.Select(line => new List<RichSeg>(line)).ToList(),
                Heights = layout.LineHeights,
                Width = layout.Lines.Count == 0 ? 0D : layout.Lines.Max(line => MeasureRichLineWidth(line, currentOpts))
            };
        }

        /// <summary>Paints the same resolved marker runs in both ordinary and row-column lists.</summary>
        private bool WriteListMarkerFallback(ListMarkerFallbackLayout? marker, double x, double width,
            double baseline, double size, double leading, PdfAlign align, global::OfficeIMO.OfficeColor? color, int? markedContentId) {
            if (marker == null) return false;
            WriteRichParagraph(sb, new RichParagraphBlock(marker.Runs, align, color), marker.Lines,
                marker.Heights, currentOpts, baseline, size, leading, currentPage!.Annotations, x, width,
                structureType: "Lbl", markedContentId: markedContentId, structurePage: currentPage);
            MarkRichFonts(marker.Runs);
            return true;
        }
    }
}
