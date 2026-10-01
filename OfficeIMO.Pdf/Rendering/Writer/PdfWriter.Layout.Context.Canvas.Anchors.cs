namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly List<(uint ZOrder, string Content)> behindTextCanvases = new();
        private double? paragraphCanvasTop;

        private void RenderParagraphCanvas(PdfCanvasBlock canvas, double top) {
            double? previous = paragraphCanvasTop;
            paragraphCanvasTop = top;
            try { RenderCanvasBlock(canvas); }
            finally { paragraphCanvasTop = previous; }
        }

        private void RenderBehindTextCanvas(PdfCanvasBehindTextItem item) {
            int start = sb.Length;
            RenderCanvasBlock(new PdfCanvasBlock(item.Items));
            string content = sb.ToString(start, sb.Length - start);
            sb.Length = start;
            behindTextCanvases.Add((item.ZOrder, content));
        }
    }
}
