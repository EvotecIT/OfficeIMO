namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>Writes a literal fallback and retains native theme-slot identity for semantic colors.</summary>
    private void SetRunColor(OfficeColor color) {
        var native = WordRunColor.Create(color, _document);
        // This property owns run/SDT/hyperlink selection and clears previous theme transforms.
        RunColorValue = native.Val!.Value!;
        ScopedRunProperties!.Color = native;
        UpdatePageNumberFormatting(native);
    }
}
