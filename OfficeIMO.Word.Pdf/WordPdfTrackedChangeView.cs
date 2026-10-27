namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Selects how tracked insertions and deletions appear in PDF output.
    /// </summary>
    public enum WordPdfTrackedChangeView {
        /// <summary>Shows the document as if every tracked change were accepted: insertions are current text and deletions are omitted.</summary>
        Final,

        /// <summary>Shows insertions underlined and deletions struck through in a revision colour, as Word prints all markup.</summary>
        Markup
    }
}
