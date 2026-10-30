namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Reports only admitted cell content; ordinary page continuations are not clipping.</summary>
        private void ReportFlowTableCellDiagnostics(PdfTableStyle style, TableCellLayout cell,
            TableCellTextLayout lines, int row, int startLine, int requestedLines, int visibleLines,
            double x, double top, double cellWidth, double cellHeight, double leading,
            double innerWidth, double verticalOffset, bool includeObjects, bool assessPageBounds = true) {
            if (!currentOpts.HasDiagnosticsReport || _diagnosticPaintHidden || !_canvasEffectToPage.TryInvert(out _)) return;
            int expectedLines = Math.Max(0, Math.Min(requestedLines, lines.LineCount - startLine));
            if (cell.TextRotation == 0 && visibleLines < expectedLines)
                ReportOmittedInlineImagePaint(cell, lines, row, startLine + visibleLines, startLine + expectedLines, x, top, cellWidth, cellHeight);
            if (cell.Viewport != null || cell.TextRotation != 0) return;
            if (visibleLines < expectedLines && !string.IsNullOrWhiteSpace(cell.Text)) {
                ReportAuthoredLayoutDiagnostic("table-cell-content-clipped", "PdfTableCell[" + row + "," + cell.Column + "]",
                    "Cell text exceeded the authored cell height and was clipped or omitted.",
                    PdfLayoutDiagnosticKind.ClippedContent, x, top, cellWidth, cellHeight);
            }

            // Plain no-wrap cells retain the regular clip. Paragraph-owned horizontal
            // frames can expand it deliberately and are outside this bounded check.
            if (cell.NoWrap && lines.LineXOffsets == null) {
                TableCellContentFrame clip = GetTableCellInkClipFrame(style, row, cell.Column, x, top, cellWidth, cellHeight);
                double bleed = style.ClipTextToCellBounds ? 0D : TableCellClipBleed;
                double textX = x + GetTableCellPaddingLeft(style, row, cell.Column);
                for (int line = startLine; line < Math.Min(lines.Lines.Count, startLine + visibleLines); line++) {
                    double effectiveAdvance = MeasureRichLineWidth(lines.Lines[line]);
                    double availablePaintWidth = clip.Left + clip.Width + bleed - textX;
                    if (effectiveAdvance <= availablePaintWidth + .001D) continue;
                    ReportAuthoredLayoutDiagnostic("table-cell-content-clipped", "PdfTableCell[" + row + "," + cell.Column + "]",
                        FormattableString.Invariant($"No-wrap cell text exceeded the authored cell width and was clipped. Effective advance {effectiveAdvance:0.###} pt exceeds available paint width {availablePaintWidth:0.###} pt."),
                        PdfLayoutDiagnosticKind.ClippedContent, x, top, cellWidth, cellHeight);
                    break;
                }
            }

            // Image target/stack dimensions do not establish painted bounds after fit,
            // crop, clipping or rotation. Leave those to the image geometry owner.
            if (!includeObjects || startLine != 0 || cell.Images.Count > 0 ||
                CanRenderTableCellCheckBoxInline(cell, lines, startLine, visibleLines)) return;
            double objectsHeight = MeasureTableCellObjectStackHeight(cell, innerWidth);
            if (objectsHeight <= 0D) return;
            double contentHeight = MeasureTableCellContentHeight(cell, lines, startLine, visibleLines, leading, innerWidth);
            // Objects use the admitted text, not the omitted lines, and are preserved outside
            // ordinary flow cells. Overflow into padding alone does not escape the cell.
            double bottom = top - GetTableCellPaddingTop(style, row, cell.Column) - verticalOffset - contentHeight;
            if (bottom < top - cellHeight - .001D) {
                if (assessPageBounds && _canvasEffectToPage.Equals(OfficeIMO.Drawing.OfficeTransform.Identity) && bottom < -.001D) {
                    ReportAuthoredLayoutDiagnostic("table-cell-object-page-bounds-clipped", "PdfTableCell[" + row + "," + cell.Column + "]",
                        "Cell objects extend below the physical page and were clipped by the page bounds.",
                        PdfLayoutDiagnosticKind.ClippedContent, x, top, cellWidth, top - bottom);
                } else {
                    ReportAuthoredLayoutDiagnostic("table-cell-object-overflow", "PdfTableCell[" + row + "," + cell.Column + "]",
                        "Cell objects extend below the authored cell bounds and were preserved.",
                        PdfLayoutDiagnosticKind.Overflow, x, top, cellWidth, cellHeight,
                        PdfConversionWarningSeverity.Information);
                }
            }
        }

        /// <summary>Proves full-line loss from the composed font envelope; advances alone do not establish clipping.</summary>
        private void ReportFullyClippedFrameLine(string code, string source,
            IReadOnlyList<List<RichSeg>> lines, IReadOnlyList<double> heights,
            double firstBaseline, double fontSize, double leading, PdfStandardFont font,
            double clipBottom, double clipHeight, double x, double top, double itemWidth, double itemHeight) {
            if (!currentOpts.HasDiagnosticsReport || _diagnosticPaintHidden || !_canvasEffectToPage.TryInvert(out _)) return;
            double baseline = firstBaseline;
            for (int index = 0; index < lines.Count; index++) {
                List<RichSeg> line = lines[index];
                if (line.Any(segment => !segment.IsLineBoxStrut &&
                    (segment.InlineElement is PdfInlineBox box
                        ? box.Background.HasValue || box.BorderColor.HasValue && box.BorderWidth.ToPoints() > 0D
                        : segment.InlineElement == null && !string.IsNullOrWhiteSpace(segment.Text)))) {
                    double inkBaseline = AdjustRichLineBaseline(baseline, line, currentOpts, fontSize, font);
                    GetRichLineInkMetrics(line, currentOpts, out double ascent, out double descent);
                    if (inkBaseline + ascent < clipBottom - .001D ||
                        inkBaseline - descent > clipBottom + clipHeight + .001D) {
                        ReportAuthoredLayoutDiagnostic(code, source,
                            "A visible composed text line lies wholly outside the authored text frame and was clipped.",
                            PdfLayoutDiagnosticKind.ClippedContent, x, top, itemWidth, itemHeight);
                        return;
                    }
                }
                baseline -= GetRichLineHeight(heights, index, leading);
            }
        }

        /// <summary>Keeps occurrence identity in the existing report's deduplication key and uses top-origin bounds.</summary>
        private void ReportAuthoredLayoutDiagnostic(string code, string source, string message,
            PdfLayoutDiagnosticKind kind, double x, double top, double itemWidth, double itemHeight,
            PdfConversionWarningSeverity severity = PdfConversionWarningSeverity.Warning) {
            if (!currentOpts.HasDiagnosticsReport) return;
            double yFromTop = currentOpts.PageHeight.ToPoints() - top;
            string occurrence = FormattableString.Invariant($"{source} on page {pages.Count + 1} at {x:0.###},{yFromTop:0.###}");
            currentOpts.AddLayoutDiagnostic(code, occurrence, message, kind, severity,
                x, yFromTop, itemWidth, itemHeight);
        }
    }
}
