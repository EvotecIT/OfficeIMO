namespace OfficeIMO.Word.GoogleDocs {
    public enum GoogleDocsImportMode {
        DriveExport = 0,
        Native = 1,
    }

    public enum GoogleDocsImportTabMode {
        FirstTab = 0,
        SelectedTab = 1,
        FlattenWithHeadings = 2,
    }

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

    public sealed class GoogleDocsImportOptions {
        public const long DefaultMaxResponseBytes = 64L * 1024L * 1024L;
        public GoogleDocsImportMode Mode { get; set; } = GoogleDocsImportMode.DriveExport;
        public GoogleDocsImportTabMode TabMode { get; set; } = GoogleDocsImportTabMode.FlattenWithHeadings;
        public string? TabId { get; set; }
        public GoogleDocsSuggestionsMode Suggestions { get; set; } = GoogleDocsSuggestionsMode.Accepted;
        public WordLoadOptions LoadOptions { get; set; } = new WordLoadOptions();
        public IProgress<OfficeIMO.GoogleWorkspace.Drive.GoogleDriveTransferProgress>? Progress { get; set; }
        public long MaxResponseBytes { get; set; } = DefaultMaxResponseBytes;
        public int MaxTabs { get; set; } = 100;
        public int MaxStructuralElements { get; set; } = 100_000;
        public int MaxTableCells { get; set; } = 1_000_000;
        public long MaxTextCharacters { get; set; } = 10_000_000L;
    }

    public sealed class GoogleDocsImportResult {
        public GoogleDocsImportResult(WordDocument document, GoogleDocumentReference source, OfficeIMO.GoogleWorkspace.TranslationReport report) {
            Document = document ?? throw new ArgumentNullException(nameof(document));
            Source = source ?? throw new ArgumentNullException(nameof(source));
            Report = report ?? throw new ArgumentNullException(nameof(report));
        }
        public WordDocument Document { get; }
        public GoogleDocumentReference Source { get; }
        public OfficeIMO.GoogleWorkspace.TranslationReport Report { get; }
    }
}
