namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Central list of Google Workspace OAuth scopes used by OfficeIMO extension packages.
    /// </summary>
    public static class GoogleWorkspaceScopeCatalog {
        /// <summary>Full read and write access to all files in Google Drive.</summary>
        public const string Drive = "https://www.googleapis.com/auth/drive";
        /// <summary>Access to files created or explicitly opened with the application.</summary>
        public const string DriveFile = "https://www.googleapis.com/auth/drive.file";
        /// <summary>Read-only access to all files in Google Drive.</summary>
        public const string DriveReadonly = "https://www.googleapis.com/auth/drive.readonly";
        /// <summary>Read and write access to Google Drive file metadata.</summary>
        public const string DriveMetadata = "https://www.googleapis.com/auth/drive.metadata";
        /// <summary>Read-only access to Google Drive file metadata.</summary>
        public const string DriveMetadataReadonly = "https://www.googleapis.com/auth/drive.metadata.readonly";
        /// <summary>Read and write access to Google Docs documents.</summary>
        public const string Documents = "https://www.googleapis.com/auth/documents";
        /// <summary>Read-only access to Google Docs documents.</summary>
        public const string DocumentsReadonly = "https://www.googleapis.com/auth/documents.readonly";
        /// <summary>Read and write access to Google Sheets spreadsheets.</summary>
        public const string Spreadsheets = "https://www.googleapis.com/auth/spreadsheets";
        /// <summary>Read-only access to Google Sheets spreadsheets.</summary>
        public const string SpreadsheetsReadonly = "https://www.googleapis.com/auth/spreadsheets.readonly";
        /// <summary>Read and write access to Google Slides presentations.</summary>
        public const string Presentations = "https://www.googleapis.com/auth/presentations";
        /// <summary>Read-only access to Google Slides presentations.</summary>
        public const string PresentationsReadonly = "https://www.googleapis.com/auth/presentations.readonly";

        /// <summary>Gets the least-privilege scopes used to create and edit Google Docs files.</summary>
        public static IReadOnlyList<string> DocsAuthoring { get; } = Array.AsReadOnly(new[] {
            DriveFile,
            Documents
        });

        /// <summary>Gets the least-privilege scopes used to create and edit Google Sheets files.</summary>
        public static IReadOnlyList<string> SheetsAuthoring { get; } = Array.AsReadOnly(new[] {
            DriveFile,
            Spreadsheets
        });

        /// <summary>Gets the least-privilege scopes used to create and edit Google Slides files.</summary>
        public static IReadOnlyList<string> SlidesAuthoring { get; } = Array.AsReadOnly(new[] {
            DriveFile,
            Presentations
        });

        /// <summary>Gets the combined least-privilege scopes for Docs, Sheets, and Slides authoring.</summary>
        public static IReadOnlyList<string> WorkspaceAuthoring { get; } = Array.AsReadOnly(new[] {
            DriveFile,
            Documents,
            Spreadsheets,
            Presentations
        });
    }
}
