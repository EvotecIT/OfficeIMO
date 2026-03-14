using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// Planning-time options for Word to Google Docs export.
    /// </summary>
    public sealed class GoogleDocsSaveOptions {
        public GoogleDriveFileLocation Location { get; set; } = new GoogleDriveFileLocation();
        public string? Title { get; set; }
        public bool FlattenFloatingContent { get; set; } = true;
        public bool RasterizeWordCharts { get; set; } = true;
        public bool PreserveCommentsViaDriveApi { get; set; }
        public bool IncludeHeadersAndFooters { get; set; } = true;
        public bool IncludeFootnotes { get; set; } = true;
        public bool IncludeBookmarksAsNamedRanges { get; set; } = true;
    }
}
