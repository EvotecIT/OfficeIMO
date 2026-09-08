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
        for (int index = 0; index < elements.Count; index++) {
            if (OfficeTextElements.ResolveBaseDirection(elements[index]) != OfficeTextDirection.RightToLeft) {
                return text;
            }
        }

        return string.Concat(elements.Reverse());
    }
}
