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

        internal bool HasVisibleDrawingText => _layout?.HasVisibleDrawingText ?? false;
        internal double DrawingAnchorX => _layout?.DrawingAnchorX ?? 0D;
        internal double DrawingAnchorY => _layout?.DrawingAnchorY ?? 0D;

        /// <summary>Shares one decoded replacement and charge between logical emission and drawing metadata.</summary>
        internal string ReadAndChargeActualText(TextOutputBudget budget) {
            L._decodedActualText ??= DecodeActualText(budget);
            if (!L._actualTextCharged) {
                budget.ChargeActualText(L._decodedActualText.Length);
                L._actualTextCharged = true;
            }
            return L._decodedActualText;
        }

        internal void BeginDrawingScope(PdfContentOrderKey? start, MarkedContentState? parent,
            bool hidden, double anchorX, double anchorY) {
            L._drawingStart = start; L._drawingParent = parent;
            L._drawingHidden = hidden || parent?.L._drawingHidden == true;
            L.DrawingAnchorX = anchorX; L.DrawingAnchorY = anchorY;
        }

        internal void ObserveDrawingPaint() {
            for (MarkedContentState? owner = this; owner != null; owner = owner.L._drawingParent)
                owner.L._hasDrawingPaint = true;
        }

        internal void ObserveDrawingText(bool visible, bool semanticAnchor, double x, double y) {
            for (MarkedContentState? owner = this; owner != null; owner = owner.L._drawingParent) {
                owner.L.HasVisibleDrawingText |= visible;
                if (semanticAnchor && !owner.L._hasDrawingAnchor) {
                    owner.L._hasDrawingAnchor = true;
                    owner.L.DrawingAnchorX = x; owner.L.DrawingAnchorY = y;
                }
            }
        }

        internal void CompleteDrawingScope(PdfContentOrderKey? end, TextOutputBudget budget,
            Action<DrawingActualTextScope>? completed) {
            if (_layout is null) return;
            if (completed == null || !HasActualText || L._drawingHidden || !L._hasDrawingPaint ||
                HasVisibleDrawingText || L._drawingStart == null || end == null) return;
            string text = ReadAndChargeActualText(budget);
            // Every retained scope consumes at least one character of the operation's existing budget.
            if (text.Length > 0) completed(new DrawingActualTextScope(this, text, L._drawingStart, end));
        }
    }

    private static bool IsDrawingPaintOperator(string name) => name is
        "S" or "s" or "f" or "F" or "f*" or "B" or "B*" or "b" or "b*" or "Do" or "sh" or "BI";
}
