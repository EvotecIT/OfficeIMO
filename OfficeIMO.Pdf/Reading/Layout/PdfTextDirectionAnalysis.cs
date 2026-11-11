using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfTextDirectionAnalysis {
    internal static T[] RestoreLogicalFragmentOrder<T>(IReadOnlyList<T> visualFragments,
        Func<T, string> getText, PdfReadingDirection direction, System.Threading.CancellationToken token = default) =>
        OfficeBidiTextResolver.RestoreLogicalFragmentOrder(visualFragments, getText,
            direction == PdfReadingDirection.RightToLeft ? OfficeTextDirection.RightToLeft : OfficeTextDirection.LeftToRight, token);

    internal static PdfReadingDirection Resolve(
        PdfReadingDirection requested,
        IEnumerable<string> textInSourceOrder) {
        if (requested != PdfReadingDirection.Auto) return requested;
        foreach (string text in textInSourceOrder) {
            OfficeTextDirection direction = OfficeTextElements.ResolveBaseDirection(text);
            if (direction == OfficeTextDirection.RightToLeft) return PdfReadingDirection.RightToLeft;
            if (direction == OfficeTextDirection.LeftToRight) return PdfReadingDirection.LeftToRight;
        }
        return PdfReadingDirection.LeftToRight;
    }

    /// <summary>
    /// Resolves the direction of the first strongly directional span in source order
    /// (content order, then paint order, then list position) without sorting the spans.
    /// </summary>
    internal static PdfReadingDirection ResolveInSourceOrder(
        PdfReadingDirection requested,
        IReadOnlyList<PdfTextSpan> spans) {
        if (requested != PdfReadingDirection.Auto) return requested;
        PdfTextSpan? first = null;
        OfficeTextDirection firstDirection = OfficeTextDirection.Auto;
        for (int index = 0; index < spans.Count; index++) {
            PdfTextSpan span = spans[index];
            if (first != null && CompareSourceOrder(span, first) >= 0) continue;
            OfficeTextDirection direction = OfficeTextElements.ResolveBaseDirection(span.Text);
            if (direction is not (OfficeTextDirection.RightToLeft or OfficeTextDirection.LeftToRight)) continue;
            first = span;
            firstDirection = direction;
        }
        return firstDirection == OfficeTextDirection.RightToLeft
            ? PdfReadingDirection.RightToLeft
            : PdfReadingDirection.LeftToRight;
    }

    /// <summary>Orders spans by content order key (missing keys first), then paint order.</summary>
    internal static int CompareSourceOrder(PdfTextSpan left, PdfTextSpan right) {
        int key = Comparer<PdfContentOrderKey?>.Default.Compare(left.ContentOrderKey, right.ContentOrderKey);
        return key != 0 ? key : left.PaintOrder.CompareTo(right.PaintOrder);
    }

    internal static string RestoreLogicalOrderFromGlyphPaintSequence(
        string text,
        bool glyphSequenceProgressesLeftToRight) {
        if (!glyphSequenceProgressesLeftToRight || text.Length < 2) return text;
        // The first element is never interior, so a run can reverse only when it starts right-to-left.
        int firstScalar = char.IsSurrogatePair(text, 0) ? char.ConvertToUtf32(text, 0) : text[0];
        if (OfficeUnicodeTextData.GetBidiClass(firstScalar) is not (OfficeBidiClass.R or OfficeBidiClass.AL)) return text;

        IReadOnlyList<string> elements = OfficeTextElements.Split(text);
        if (elements.Count < 2) return text;
        // Right-to-left words separated by spaces or punctuation reverse as a whole. Left-to-right
        // letters and digits, and neutrals at either edge (which may belong to the surrounding
        // context), make the logical order ambiguous, so such runs are left unchanged.
        for (int index = 0; index < elements.Count; index++) {
            string element = elements[index];
            if (element.Length == 0) return text;
            int scalar = char.IsSurrogatePair(element, 0) ? char.ConvertToUtf32(element, 0) : element[0];
            OfficeBidiClass bidiClass = OfficeUnicodeTextData.GetBidiClass(scalar);
            if (bidiClass is OfficeBidiClass.R or OfficeBidiClass.AL) continue;
            bool interior = index > 0 && index < elements.Count - 1;
            if (!interior || bidiClass is not (OfficeBidiClass.WS or OfficeBidiClass.CS or OfficeBidiClass.ON or OfficeBidiClass.ES or OfficeBidiClass.ET or OfficeBidiClass.NSM or OfficeBidiClass.BN)) return text;
        }

        return string.Concat(elements.Reverse());
    }
}
