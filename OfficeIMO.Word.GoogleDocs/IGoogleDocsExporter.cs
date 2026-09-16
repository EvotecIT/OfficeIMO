namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// Planning, compilation, and export contract for Word to Google Docs implementations.
    /// </summary>
    public interface IGoogleDocsExporter {
        /// <summary>Reports source features and fidelity risks without contacting Google.</summary>
        GoogleDocsTranslationPlan BuildPlan(WordDocument document, GoogleDocsSaveOptions? options = null);
        /// <summary>Compiles a Word snapshot into provider-neutral Docs requests without contacting Google.</summary>
        GoogleDocsBatch BuildBatch(WordDocument document, GoogleDocsSaveOptions? options = null);
        /// <summary>Creates or replaces a Google document through the supplied session.</summary>
        Task<GoogleDocumentReference> ExportAsync(
            WordDocument document,
            GoogleWorkspace.GoogleWorkspaceSession session,
            GoogleDocsSaveOptions? options = null,
            CancellationToken cancellationToken = default);
    }
}
