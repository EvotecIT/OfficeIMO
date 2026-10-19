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

    internal static string RestoreLogicalOrderFromGlyphPaintSequence(
        string text,
        bool glyphSequenceProgressesLeftToRight) {
        if (!glyphSequenceProgressesLeftToRight || text.Length < 2) return text;

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
