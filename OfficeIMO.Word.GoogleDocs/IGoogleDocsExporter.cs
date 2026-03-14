namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// Future export contract for Word to Google Docs implementations.
    /// </summary>
    public interface IGoogleDocsExporter {
        GoogleDocsTranslationPlan BuildPlan(WordDocument document, GoogleDocsSaveOptions? options = null);
        GoogleDocsBatch BuildBatch(WordDocument document, GoogleDocsSaveOptions? options = null);
        Task<GoogleDocumentReference> ExportAsync(
            WordDocument document,
            GoogleWorkspace.GoogleWorkspaceSession session,
            GoogleDocsSaveOptions? options = null,
            CancellationToken cancellationToken = default);
    }
}
