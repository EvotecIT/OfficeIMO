namespace OfficeIMO.Pdf;

internal static partial class TextContentParser {
    private static bool HasLogicalBoundsHint(object? propertyObject) =>
        propertyObject is PdfContentDictionary dictionary &&
        dictionary.Items.TryGetValue("OfficeIMOLogicalBounds", out object? value) && value is bool flag && flag;

    internal sealed partial class MarkedContentState {
        internal bool PreserveAnchorGeometry { get; }
        private List<PdfTextSpan>? _invisibleAnchorOwner;
        private int _invisibleAnchorIndex;

        internal void RememberInvisibleAnchor(List<PdfTextSpan> owner, int index) {
            if (PreserveAnchorGeometry) return;
            _invisibleAnchorOwner = owner;
            _invisibleAnchorIndex = index;
        }

        internal bool TryGetInvisibleAnchorText(List<PdfTextSpan> owner, out string? text) {
            // Form parsing shares replacement ownership, but its span list has a
            // separate lifetime. Never replace an already returned parent stream.
            if (!ReferenceEquals(owner, _invisibleAnchorOwner)) {
                text = null;
                return false;
            }
            PdfTextSpan anchor = owner[_invisibleAnchorIndex];
            text = anchor.SourceActualText ?? anchor.Text;
            return true;
        }

        internal void ReplaceInvisibleAnchor(List<PdfTextSpan> owner, PdfTextSpan paint) {
            if (!ReferenceEquals(owner, _invisibleAnchorOwner))
                throw new InvalidOperationException("ActualText geometry must belong to the current content stream.");
            // Geometry promotion spans multiple text objects and keeps a separate
            // semantic carrier. Removing just the paint object cannot safely edit
            // that replacement or preserve graphics state consumed later.
            owner[_invisibleAnchorIndex] = paint.WithCanRestamp(false);
            _invisibleAnchorOwner = null;
        }
    }
}
