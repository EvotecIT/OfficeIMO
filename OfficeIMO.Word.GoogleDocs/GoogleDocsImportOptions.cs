using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>How native import selects or combines document tabs.</summary>
    public enum GoogleDocsImportTabMode {
        /// <summary>Import only the first tab.</summary>
        FirstTab = 0,
        /// <summary>Import the tab identified by <see cref="GoogleDocsImportOptions.TabId"/>.</summary>
        SelectedTab = 1,
        /// <summary>Combine tab content into one Word document with distinguishing headings.</summary>
        FlattenWithHeadings = 2,
    }

    /// <summary>Suggestions view requested from the native Google Docs API.</summary>
    public enum GoogleDocsSuggestionsMode {
        /// <summary>Uses the API default for the current principal.</summary>
        Default = 0,
        /// <summary>Returns a preview with suggestions accepted.</summary>
        Accepted = 1,
        /// <summary>Returns suggestions inline with the document content.</summary>
        Inline = 2,
        /// <summary>Returns a preview with suggestions rejected.</summary>
        Rejected = 3,
    }

    /// <summary>Import mode, tab/suggestion selection, and bounded projection limits.</summary>
    public sealed class GoogleDocsImportOptions {
        /// <summary>Default response-size limit for Drive-exported DOCX or native Docs JSON, in bytes.</summary>
        public const long DefaultMaxResponseBytes = 64L * 1024L * 1024L;
        /// <summary>Gets or sets Drive-export or native Docs import; defaults to Drive export.</summary>
        public GoogleWorkspaceImportMode Mode { get; set; } = GoogleWorkspaceImportMode.DriveExport;
        /// <summary>Gets or sets native tab handling; defaults to flattening with headings.</summary>
        public GoogleDocsImportTabMode TabMode { get; set; } = GoogleDocsImportTabMode.FlattenWithHeadings;
        /// <summary>Gets or sets the tab identifier required by selected-tab native import.</summary>
        public string? TabId { get; set; }
        /// <summary>Gets or sets how the Docs API presents suggestions; defaults to an accepted preview.</summary>
        public GoogleDocsSuggestionsMode Suggestions { get; set; } = GoogleDocsSuggestionsMode.Accepted;
        /// <summary>Gets or sets options used to load Drive-exported DOCX.</summary>
        public WordLoadOptions LoadOptions { get; set; } = new WordLoadOptions();
        /// <summary>Gets or sets an optional observer for Drive-export transfer progress.</summary>
        public IProgress<OfficeIMO.GoogleWorkspace.Drive.GoogleDriveTransferProgress>? Progress { get; set; }
        /// <summary>Gets or sets the positive byte limit for Drive-exported DOCX or native API response.</summary>
        public long MaxResponseBytes { get; set; } = DefaultMaxResponseBytes;
        /// <summary>Gets or sets the positive maximum number of tabs accepted by native import.</summary>
        public int MaxTabs { get; set; } = 100;
        /// <summary>Gets or sets the positive structural-element limit for body content across all returned tabs, including unselected tabs.</summary>
        /// <remarks>Header, footer, and footnote segments are not counted here; <see cref="MaxResponseBytes"/> bounds the complete response.</remarks>
        public int MaxStructuralElements { get; set; } = 100_000;
        /// <summary>
        /// Maximum aggregate rectangular table-cell projection in body content across all returned
        /// tabs, including unselected tabs. The same ceiling independently bounds aggregate table rows
        /// so sparse tables cannot allocate unbounded document rows.
        /// </summary>
        /// <remarks>Header, footer, and footnote segments are not counted here; <see cref="MaxResponseBytes"/> bounds the complete response.</remarks>
        public int MaxTableCells { get; set; } = 1_000_000;
        /// <summary>Gets or sets the positive text-character limit for body content across all returned tabs, including unselected tabs.</summary>
        /// <remarks>Header, footer, and footnote segments are not counted here; <see cref="MaxResponseBytes"/> bounds the complete response.</remarks>
        public long MaxTextCharacters { get; set; } = 10_000_000L;
    }

    /// <summary>An imported Word document, remote reference, and fidelity notices.</summary>
    public sealed class GoogleDocsImportResult {
        /// <summary>Creates an import result from a caller-owned document and source evidence.</summary>
        public GoogleDocsImportResult(WordDocument document, GoogleDocumentReference source, OfficeIMO.GoogleWorkspace.TranslationReport report) {
            Document = document ?? throw new ArgumentNullException(nameof(document));
            Source = source ?? throw new ArgumentNullException(nameof(source));
            Report = report ?? throw new ArgumentNullException(nameof(report));
        }
        /// <summary>Gets the imported Word document, which the caller must dispose.</summary>
        public WordDocument Document { get; }
        /// <summary>Gets remote metadata; a Docs revision is observed only in native import mode, not Drive-export mode.</summary>
        public GoogleDocumentReference Source { get; }
        /// <summary>Gets fidelity and operation notices from import.</summary>
        public OfficeIMO.GoogleWorkspace.TranslationReport Report { get; }
    }
}
