namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// Planning helpers for Word to Google Docs translation.
    /// </summary>
    public static class WordGoogleDocsExtensions {
        public static GoogleDocsTranslationPlan CreateGoogleDocsTranslationPlan(
            this WordDocument document,
            GoogleDocsSaveOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return GoogleDocsPlanBuilder.Build(document, options ?? new GoogleDocsSaveOptions());
        }
    }
}
