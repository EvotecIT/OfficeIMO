namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly List<(long ZOrder, string Content)> behindTextCanvases = new();
        private readonly List<(long ZOrder, string Content)> inFrontOfTextCanvases = new();
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

        private void RenderInFrontOfTextCanvas(PdfCanvasInFrontOfTextItem item) {
            int start = sb.Length;
            RenderCanvasBlock(new PdfCanvasBlock(item.Items));
            string content = sb.ToString(start, sb.Length - start);
            sb.Length = start;
            inFrontOfTextCanvases.Add((item.ZOrder, content));
        }

        /// <summary>Places deferred behind-text paint under the story's content and in-front paint over it.</summary>
        private void ApplyDeferredTextLayers() {
            if (behindTextCanvases.Count > 0) {
                var behind = new System.Text.StringBuilder();
                foreach (var canvas in behindTextCanvases.OrderBy(item => item.ZOrder)) behind.Append(canvas.Content);
                sb.Insert(0, behind.ToString());
                behindTextCanvases.Clear();
            }
            if (inFrontOfTextCanvases.Count > 0) {
                foreach (var canvas in inFrontOfTextCanvases.OrderBy(item => item.ZOrder)) sb.Append(canvas.Content);
                inFrontOfTextCanvases.Clear();
            }
        }
    }
}
