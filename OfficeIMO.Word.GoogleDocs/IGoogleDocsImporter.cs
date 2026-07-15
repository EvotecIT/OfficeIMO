namespace OfficeIMO.Word.GoogleDocs {
    public interface IGoogleDocsImporter {
        Task<GoogleDocsImportResult> ImportAsync(
            string documentId,
            GoogleWorkspace.GoogleWorkspaceSession session,
            GoogleDocsImportOptions? options = null,
            CancellationToken cancellationToken = default);
    }
}
