using System.Collections.Generic;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Retains one paragraph frame around objects that actually enter flow.</summary>
    private sealed class NativeObjectParagraphSpacing {
        private readonly INativePdfFlow _flow;
        private readonly OfficeIMO.Pdf.PdfParagraphStyle _style;
        private readonly double _minimumHeight;
        private readonly List<(Action<INativePdfFlow> Render, bool AlignToLineTop)> _objects = new();

        internal NativeObjectParagraphSpacing(INativePdfFlow flow, OfficeIMO.Pdf.PdfParagraphStyle style, double minimumHeight) {
            _flow = flow;
            _style = style;
            _minimumHeight = minimumHeight;
        }

        internal void Add(Action<INativePdfFlow> render, bool alignToLineTop) =>
            _objects.Add((render, alignToLineTop));

        internal bool Complete() {
            if (_objects.Count == 0) return false;
            // The shared panel keeps the line minimum, canvas and objects on the
            // same page or column. It adds no decoration or extra paragraph mark.
            _flow.Panel(content => {
                var inner = new NativePdfColumnFlow(content, _flow.PageSize);
                foreach (var item in _objects) item.Render(inner);
            }, new OfficeIMO.Pdf.PdfPanelStyle {
                PaddingX = 0D, PaddingY = 0D, BorderWidth = 0D,
                SpacingBefore = _style.SpacingBefore, SpacingAfter = _style.SpacingAfter ?? 0D,
                KeepTogether = true, KeepWithNext = _style.KeepWithNext,
                AnchoredCanvas = _style.AnchoredCanvas,
                MinimumContentHeight = _minimumHeight,
                // Inline objects align with the line's baseline. VML and anchored
                // shapes retain their line-top position when the minimum grows.
                AlignContentToBottom = !_objects.All(item => item.AlignToLineTop)
            });
            return true;
        }
    }

    private static void RenderNativeFlowObject(INativePdfFlow flow, NativeObjectParagraphSpacing? paragraph,
        Action<INativePdfFlow> render, bool alignToLineTop = false) {
        if (paragraph == null) render(flow);
        else paragraph.Add(render, alignToLineTop);
    }
}
