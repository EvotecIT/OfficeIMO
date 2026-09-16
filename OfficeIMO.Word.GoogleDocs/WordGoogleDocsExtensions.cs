namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// Helpers for Word to Google Docs planning and batch compilation.
    /// </summary>
    public static class WordGoogleDocsExtensions {
        /// <summary>Imports a Google document and returns a caller-owned Word document.</summary>
        public static Task<GoogleDocsImportResult> ImportGoogleDocAsync(
            this GoogleWorkspace.GoogleWorkspaceSession session,
            string documentId,
            GoogleDocsImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            return new GoogleDocsImporter().ImportAsync(documentId, session, options, cancellationToken);
        }
        private static readonly IGoogleDocsExporter DefaultExporter = new GoogleDocsExporter();

        /// <summary>Reports Word features and fidelity risks without contacting Google.</summary>
        public static GoogleDocsTranslationPlan BuildGoogleDocsPlan(
            this WordDocument document,
            GoogleDocsSaveOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return DefaultExporter.BuildPlan(document, options);
        }

        /// <summary>Compiles Word content into provider-neutral Docs requests without contacting Google.</summary>
        public static GoogleDocsBatch BuildGoogleDocsBatch(
            this WordDocument document,
            GoogleDocsSaveOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return DefaultExporter.BuildBatch(document, options);
        }

        /// <summary>Creates or replaces a Google document using the supplied session.</summary>
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
