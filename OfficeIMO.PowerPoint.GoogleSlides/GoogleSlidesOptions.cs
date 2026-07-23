using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.PowerPoint.GoogleSlides {
    public enum GoogleSlidesComplexSlideMode {
        PreferNativeAndReport = 0,
        RasterizeComplexSlides = 1,
    }

    public enum GoogleSlidesRevisionConflictMode {
        RequireRevision = 0,
        OverwriteLatest = 1,
    }

    public sealed class GoogleSlidesReplaceOptions {
        public GoogleSlidesRevisionConflictMode ConflictMode { get; set; } = GoogleSlidesRevisionConflictMode.RequireRevision;
        public string? ExpectedRevisionId { get; set; }
    }

    public sealed class GoogleSlidesSaveOptions {
        public string? Title { get; set; }
        public GoogleDriveFileLocation Location { get; set; } = new GoogleDriveFileLocation();
        public string? TemplatePresentationId { get; set; }
        public GoogleSlidesComplexSlideMode ComplexSlides { get; set; } = GoogleSlidesComplexSlideMode.RasterizeComplexSlides;
        public GoogleSlidesReplaceOptions Replace { get; set; } = new GoogleSlidesReplaceOptions();
        public GoogleWorkspaceFidelityPolicy FidelityPolicy { get; set; } = new GoogleWorkspaceFidelityPolicy();
    }

    public enum GoogleSlidesImportMode {
        DriveExport = 0,
        Native = 1,
    }

    public sealed class GoogleSlidesImportOptions {
        public GoogleSlidesImportMode Mode { get; set; } = GoogleSlidesImportMode.DriveExport;
        public PowerPointLoadOptions LoadOptions { get; set; } = new PowerPointLoadOptions();
        public IProgress<OfficeIMO.GoogleWorkspace.Drive.GoogleDriveTransferProgress>? Progress { get; set; }

        /// <summary>
        /// Maximum bytes accepted for each Google-hosted image during native import.
        /// </summary>
        public long MaxImageBytes { get; set; } = 50L * 1024 * 1024;
    }
}
