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
        Default = 0,
        Accepted = 1,
        Inline = 2,
    }

    public sealed class GoogleDocsImportOptions {
        public GoogleDocsImportMode Mode { get; set; } = GoogleDocsImportMode.DriveExport;
        public GoogleDocsImportTabMode TabMode { get; set; } = GoogleDocsImportTabMode.FlattenWithHeadings;
        public string? TabId { get; set; }
        public GoogleDocsSuggestionsMode Suggestions { get; set; } = GoogleDocsSuggestionsMode.Accepted;
        public WordLoadOptions LoadOptions { get; set; } = new WordLoadOptions();
        public IProgress<OfficeIMO.GoogleWorkspace.Drive.GoogleDriveTransferProgress>? Progress { get; set; }
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
