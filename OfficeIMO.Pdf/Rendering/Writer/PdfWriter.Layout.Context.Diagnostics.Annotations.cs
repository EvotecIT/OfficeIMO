namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Checks the final rectangle used by appearance generation, after canvas annotation transforms and clipping.</summary>
        private void ReportFreeTextAnnotationDiagnostics() {
            if (!currentOpts.HasDiagnosticsReport || currentPage == null) return;
            for (int annotationIndex = 0; annotationIndex < currentPage.FreeTextAnnotations.Count; annotationIndex++) {
                FreeTextAnnotation annotation = currentPage.FreeTextAnnotations[annotationIndex];
                double width = annotation.X2 - annotation.X1, height = annotation.Y2 - annotation.Y1;
                double innerWidth = Math.Max(0D, width - 2D * annotation.Padding);
                double innerHeight = Math.Max(0D, height - 2D * annotation.Padding);
                double leading = annotation.LineHeight ?? annotation.FontSize * 1.2D;
                var runs = new[] { PdfTextRun.Normal(annotation.Contents, annotation.TextColor, annotation.FontSize, font: PdfStandardFont.Helvetica) };
                // Wrapping omitted lines must not enlarge the serialized font subset.
                var measurementOptions = DiagnosticTextOptions;
                var wrapped = WrapRichRunsCore(runs, innerWidth, annotation.FontSize, PdfStandardFont.Helvetica,
                    leading, null, DefaultParagraphTabStopWidth, measurementOptions);
                int limit = innerHeight > 0D ? Math.Max(1, (int)Math.Floor(innerHeight / leading)) : 0;
                string source = "PdfFreeTextAnnotation[" + annotationIndex + "]";
                double baseline = height - annotation.Padding - annotation.FontSize;
                bool omitted = false;
                for (int line = 0; line < wrapped.Lines.Count; line++, baseline -= leading) {
                    var segments = wrapped.Lines[line];
                    bool admitted = line < limit && baseline >= annotation.Padding - annotation.FontSize * .25D;
                    if (!admitted) {
                        // Empty line slots and unused leading are not lost paint.
                        omitted |= segments.Any(segment => !string.IsNullOrWhiteSpace(segment.Text));
                        continue;
                    }
                    if (!CanEncodeFreeTextAppearanceLine(segments, measurementOptions)) continue; // Encoding has its own diagnostic owner.
                    var paragraph = new RichParagraphBlock(runs, annotation.TextAlign, annotation.TextColor);
                    // Appearance baselines do not use paragraph line-box correction.
                    var lines = new List<List<RichSeg>> { new List<RichSeg>(segments) };
                    ReportRichPaintDiagnostics("FreeTextAnnotation", source, paragraph, lines, new[] { leading },
                        annotation.Y1 + Math.Max(0D, baseline), annotation.FontSize, leading,
                        annotation.X1 + annotation.Padding, innerWidth,
                        new TableCellContentFrame(annotation.X1, annotation.Y2, width, height), baselineFont: PdfStandardFont.Helvetica);
                }
                if (omitted) ReportAuthoredLayoutDiagnostic("FreeTextAnnotationContentOmitted", source,
                    "Composed annotation text exceeded the appearance's admitted lines and was omitted.",
                    PdfLayoutDiagnosticKind.ClippedContent, annotation.X1, annotation.Y2, width, height);
            }
        }
    }
}
