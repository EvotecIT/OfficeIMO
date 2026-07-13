namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// Helpers for Word to Google Docs planning and batch compilation.
    /// </summary>
    public static class WordGoogleDocsExtensions {
        private static readonly IGoogleDocsExporter DefaultExporter = new GoogleDocsExporter();

        public static GoogleDocsTranslationPlan BuildGoogleDocsPlan(
            this WordDocument document,
            GoogleDocsSaveOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return DefaultExporter.BuildPlan(document, options);
        }

        public static GoogleDocsBatch BuildGoogleDocsBatch(
            this WordDocument document,
            GoogleDocsSaveOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return DefaultExporter.BuildBatch(document, options);
        }

        public static Task<GoogleDocumentReference> ExportToGoogleDocsAsync(
            this WordDocument document,
            GoogleWorkspace.GoogleWorkspaceSession session,
            GoogleDocsSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return DefaultExporter.ExportAsync(document, session, options, cancellationToken);
        }
    }
}
