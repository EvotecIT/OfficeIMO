namespace OfficeIMO.Pdf;

public sealed partial class PdfTextRun {
    /// <summary>One descriptor instance identifies all styled runs of the same source control.</summary>
    internal PdfInlineFormField? FormField { get; private set; }

    internal PdfTextRun WithFormField(PdfInlineFormField field) {
        Guard.NotNull(field, nameof(field));
        // The run's visual state is immutable outside its copy helpers; preserve every
        // existing run property, including inline spacers and shaped glyph metadata.
        var copy = (PdfTextRun)MemberwiseClone();
        copy.FormField = field;
        return copy;
    }

    internal PdfTextRun SnapshotForInlineAppearance() {
        if (InlineElement != null)
            throw new ArgumentException("A form glyph appearance requires a text run.");
        var copy = (PdfTextRun)MemberwiseClone();
        copy.FormField = null;
        return copy;
    }
}
