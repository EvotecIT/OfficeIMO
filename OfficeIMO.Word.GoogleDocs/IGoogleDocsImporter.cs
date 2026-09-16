namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>Imports a Google document into a caller-owned Word document.</summary>
    public interface IGoogleDocsImporter {
        /// <summary>Loads a document through Drive-exported DOCX or native Docs projection.</summary>
        Task<GoogleDocsImportResult> ImportAsync(
            string documentId,
            GoogleWorkspace.GoogleWorkspaceSession session,
            GoogleDocsImportOptions? options = null,
            CancellationToken cancellationToken = default);
    }
}
