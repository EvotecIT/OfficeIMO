using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// Result metadata for a created or updated Google Doc.
    /// </summary>
    public sealed class GoogleDocumentReference : GoogleDriveFileReference {
        public string? DocumentId { get; set; }
        public TranslationReport Report { get; set; } = new TranslationReport();
    }
}
