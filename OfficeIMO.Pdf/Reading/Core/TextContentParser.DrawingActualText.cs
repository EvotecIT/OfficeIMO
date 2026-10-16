namespace OfficeIMO.Pdf;

internal static partial class TextContentParser {
    /// <summary>A locally closed replacement scope, ordered with the paint interpreted from the same stream.</summary>
    internal sealed class DrawingActualTextScope {
        internal DrawingActualTextScope(MarkedContentState owner, string text, PdfContentOrderKey start,
            PdfContentOrderKey end) {
            Owner = owner; Text = text; Start = start; End = end;
        }

        internal MarkedContentState Owner { get; }
        internal string Text { get; }
        internal PdfContentOrderKey Start { get; }
        internal PdfContentOrderKey End { get; }
    }

    internal sealed partial class MarkedContentState {
        private string? _decodedActualText;
        private bool _actualTextCharged;
        private PdfContentOrderKey? _drawingStart;
        private MarkedContentState? _drawingParent;
        private bool _drawingHidden;
        private bool _hasDrawingPaint;
        private bool _hasDrawingAnchor;

        internal bool HasVisibleDrawingText { get; private set; }
        internal double DrawingAnchorX { get; private set; }
        internal double DrawingAnchorY { get; private set; }

        /// <summary>Shares one decoded replacement and charge between logical emission and drawing metadata.</summary>
        internal string ReadAndChargeActualText(TextOutputBudget budget) {
            _decodedActualText ??= DecodeActualText(budget);
            if (!_actualTextCharged) {
                budget.ChargeActualText(_decodedActualText.Length);
                _actualTextCharged = true;
            }
            return _decodedActualText;
        }

        internal void BeginDrawingScope(PdfContentOrderKey? start, MarkedContentState? parent,
            bool hidden, double anchorX, double anchorY) {
            _drawingStart = start; _drawingParent = parent;
            _drawingHidden = hidden || parent?._drawingHidden == true;
            DrawingAnchorX = anchorX; DrawingAnchorY = anchorY;
        }

        internal void ObserveDrawingPaint() {
            for (MarkedContentState? owner = this; owner != null; owner = owner._drawingParent)
                owner._hasDrawingPaint = true;
        }

        internal void ObserveDrawingText(bool visible, bool semanticAnchor, double x, double y) {
            for (MarkedContentState? owner = this; owner != null; owner = owner._drawingParent) {
                owner.HasVisibleDrawingText |= visible;
                if (semanticAnchor && !owner._hasDrawingAnchor) {
                    owner._hasDrawingAnchor = true;
                    owner.DrawingAnchorX = x; owner.DrawingAnchorY = y;
                }
            }
        }

        internal void CompleteDrawingScope(PdfContentOrderKey? end, TextOutputBudget budget,
            Action<DrawingActualTextScope>? completed) {
            if (completed == null || !HasActualText || _drawingHidden || !_hasDrawingPaint ||
                HasVisibleDrawingText || _drawingStart == null || end == null) return;
            string text = ReadAndChargeActualText(budget);
            // Every retained scope consumes at least one character of the operation's existing budget.
            if (text.Length > 0) completed(new DrawingActualTextScope(this, text, _drawingStart, end));
        }
    }

    private static bool IsDrawingPaintOperator(string name) => name is
        "S" or "s" or "f" or "F" or "f*" or "B" or "B*" or "b" or "b*" or "Do" or "sh" or "BI";
}
