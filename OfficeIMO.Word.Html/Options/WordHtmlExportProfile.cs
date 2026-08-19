namespace OfficeIMO.Word.Html {
    /// <summary>Named Word-to-HTML output contracts.</summary>
    public enum WordHtmlExportProfile {
        /// <summary>Readable, accessible document structure for review and publishing.</summary>
        SemanticDocument,

        /// <summary>Trusted editable round-trip output with private OfficeIMO metadata.</summary>
        DocumentRoundTrip,

        /// <summary>Print-oriented review output without claiming full Word layout parity.</summary>
        PrintReview
    }
}
