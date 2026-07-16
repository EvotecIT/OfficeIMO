namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Central list of Google Workspace OAuth scopes used by OfficeIMO extension packages.
    /// </summary>
    public static class GoogleWorkspaceScopeCatalog {
        public const string Drive = "https://www.googleapis.com/auth/drive";
        public const string DriveFile = "https://www.googleapis.com/auth/drive.file";
        public const string DriveReadonly = "https://www.googleapis.com/auth/drive.readonly";
        public const string DriveMetadata = "https://www.googleapis.com/auth/drive.metadata";
        public const string DriveMetadataReadonly = "https://www.googleapis.com/auth/drive.metadata.readonly";
        public const string Documents = "https://www.googleapis.com/auth/documents";
        public const string DocumentsReadonly = "https://www.googleapis.com/auth/documents.readonly";
        public const string Spreadsheets = "https://www.googleapis.com/auth/spreadsheets";
        public const string SpreadsheetsReadonly = "https://www.googleapis.com/auth/spreadsheets.readonly";
        public const string Presentations = "https://www.googleapis.com/auth/presentations";
        public const string PresentationsReadonly = "https://www.googleapis.com/auth/presentations.readonly";

        public static IReadOnlyList<string> DocsAuthoring { get; } = new[] {
            DriveFile,
            Documents
        };

        public static IReadOnlyList<string> SheetsAuthoring { get; } = new[] {
            DriveFile,
            Spreadsheets
        };

        public static IReadOnlyList<string> SlidesAuthoring { get; } = new[] {
            DriveFile,
            Presentations
        };

        public static IReadOnlyList<string> WorkspaceAuthoring { get; } = new[] {
            DriveFile,
            Documents,
            Spreadsheets,
            Presentations
        };
    }
}
