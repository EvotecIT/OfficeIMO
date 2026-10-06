using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        // A contiguous baseline is one PDF text run. Encoding each cluster with a
        // new text matrix causes native readers to infer separate transformed lines.
        // Explicit widths retain cluster advances without changing the text matrix.
        private bool TryRenderSearchableRun(IReadOnlyList<PdfCanvasItem> items, ref int index) {
            var first = (PdfCanvasSearchableTextItem)items[index];
            if (first.Geometry is not PdfSelectionQuad quad) return false;
            double dx = quad.BottomRight.X - quad.BottomLeft.X, dy = quad.BottomRight.Y - quad.BottomLeft.Y;
            double ux = quad.TopLeft.X - quad.BottomLeft.X, uy = quad.TopLeft.Y - quad.BottomLeft.Y;
            double width = Math.Sqrt(dx * dx + dy * dy), height = Math.Sqrt(ux * ux + uy * uy);
            if (width <= 0 || height <= 0) return false;
            double ax = dx / width, ay = dy / width;
            int last = index;
            var previous = quad;
            while (last + 1 < items.Count && items[last + 1] is PdfCanvasSearchableTextItem next && next.Geometry is PdfSelectionQuad q) {
                double nx = q.BottomRight.X - q.BottomLeft.X, ny = q.BottomRight.Y - q.BottomLeft.Y;
                double length = Math.Sqrt(nx * nx + ny * ny);
                if (next.LogicalOrder != first.LogicalOrder || length <= 0 || !Close(q.BottomLeft.X, previous.BottomRight.X) || !Close(q.BottomLeft.Y, previous.BottomRight.Y) ||
                    !Close(q.TopLeft.X - q.BottomLeft.X, ux) || !Close(q.TopLeft.Y - q.BottomLeft.Y, uy) ||
                    !Close(nx / length, ax) || !Close(ny / length, ay)) break;
                previous = q; last++;
            }
            if (last == index) return false;
            if (_suppressCanvasActualTextChildren) { index = last; return true; }
            EnsurePage();
            PdfStandardFont font = ChooseNormal(currentOpts.DefaultFont);
            PdfTextShowCommand space = EncodeBoundedLogicalTextAnchor(font, currentOpts, 1);
            bool cff = currentOpts.TryGetEmbeddedStandardOpenTypeCffFontProgram(font, out var cffProgram);
            var content = new ContentStreamBuilder(sb).SaveState().BeginText()
                .WordSpacing(0D).TextRise(0D).HorizontalTextScaling(100D).TextRenderingMode(3)
                .TextMatrix(ax, -ay, ux / height, -uy / height, quad.BottomLeft.X, currentOpts.PageHeight - quad.BottomLeft.Y);
            var logicalText = new StringBuilder();
            for (int i = index; i <= last; i++) logicalText.Append(((PdfCanvasSearchableTextItem)items[i]).Text);
            int? markedContentId = RegisterTextStructureElement("Span", _canvasStructureParentElement, logicalOrder: first.LogicalOrder);
            sb.Append("/Span << /ActualText ").Append(PdfSyntaxEscaper.TextString(logicalText.ToString()));
            if (markedContentId.HasValue) sb.Append(" /MCID ").Append(markedContentId.Value.ToString(CultureInfo.InvariantCulture));
            sb.Append(" >> BDC\n");
            SearchableFontBank? bank = null;
            var hex = new StringBuilder();
            for (int i = index; i <= last; i++) {
                var item = (PdfCanvasSearchableTextItem)items[i];
                var geometry = item.Geometry!;
                double advanceX = geometry.BottomRight.X - geometry.BottomLeft.X, advanceY = geometry.BottomRight.Y - geometry.BottomLeft.Y;
                double scalarWidth = Math.Sqrt(advanceX * advanceX + advanceY * advanceY) * 1000D / height / CountLogicalAnchorScalars(item.Text);
                foreach (var run in currentPage!.SearchableFonts.Encode(item.Text, space, cff, scalarWidth, cffProgram)) {
                    if (bank != null && bank != run.Bank) {
                        content.Font(bank.Name, height, preserveLogicalPrecision: true).ShowHexText(hex.ToString()); hex.Clear();
                    }
                    bank = run.Bank; hex.Append(run.Hex);
                }
            }
            if (bank != null) content.Font(bank.Name, height, preserveLogicalPrecision: true).ShowHexText(hex.ToString());
            sb.Append("EMC\n");
            content.EndText().RestoreState();
            MarkSimpleFont(font); pageDirty = true; index = last;
            return true;
        }

        private static bool Close(double left, double right) => Math.Abs(left - right) <= 0.0000001D;
    }
}
