using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// Result metadata for a created or updated Google Doc.
    /// </summary>
    public sealed class GoogleDocumentReference : GoogleDriveFileReference {
        /// <summary>Gets or sets the Docs document identifier.</summary>
        public string? DocumentId { get; set; }
        /// <summary>Gets or sets the Docs revision observed by native import or export; Drive-export import leaves it null.</summary>
        public string? RevisionId { get; set; }
        /// <summary>Gets or sets the observed Drive version, when available.</summary>
        public long? DriveVersion { get; set; }
        /// <summary>Gets or sets the observed modification time, when available.</summary>
        public DateTimeOffset? ModifiedTime { get; set; }
        /// <summary>Gets or sets translation notices associated with this reference.</summary>
        public TranslationReport Report { get; set; } = new TranslationReport();
    }
}
