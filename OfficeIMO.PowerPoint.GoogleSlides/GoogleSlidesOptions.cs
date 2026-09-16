using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.PowerPoint.GoogleSlides {
    /// <summary>How export handles a slide with content that lacks a dependable native Slides equivalent.</summary>
    public enum GoogleSlidesComplexSlideMode {
        /// <summary>Keep supported elements editable and report skipped unsupported content.</summary>
        PreferNativeAndReport = 0,
        /// <summary>Render the whole complex slide to an image to preserve its appearance.</summary>
        RasterizeComplexSlides = 1,
    }

    /// <summary>Revision behavior when replacing an existing Google presentation.</summary>
    public enum GoogleSlidesRevisionConflictMode {
        /// <summary>Require a previously observed revision and reject a stale replacement.</summary>
        RequireRevision = 0,
        /// <summary>Replace the latest remote content without a Slides revision write guard.</summary>
        OverwriteLatest = 1,
    }

    /// <summary>Controls optimistic-concurrency behavior for an existing presentation.</summary>
    public sealed class GoogleSlidesReplaceOptions {
        /// <summary>Gets or sets the replacement conflict mode; defaults to requiring a revision.</summary>
        public GoogleSlidesRevisionConflictMode ConflictMode { get; set; } = GoogleSlidesRevisionConflictMode.RequireRevision;
        /// <summary>Gets or sets the revision observed during an earlier read or import.</summary>
        /// <remarks>Required for an existing presentation unless <see cref="GoogleSlidesRevisionConflictMode.OverwriteLatest"/> is selected.</remarks>
        public string? ExpectedRevisionId { get; set; }
    }

    /// <summary>Destination, fidelity, and replacement choices for Slides export.</summary>
    public sealed class GoogleSlidesSaveOptions {
        /// <summary>Gets or sets the presentation title; blank uses the source title or <c>Presentation</c>.</summary>
        public string? Title { get; set; }
        /// <summary>Gets or sets Drive folder, shared-drive, and existing-file targeting information.</summary>
        public GoogleDriveFileLocation Location { get; set; } = new GoogleDriveFileLocation();
        /// <summary>Gets or sets a template presentation to copy before authoring.</summary>
        /// <remarks>When set, the exporter copies the template rather than replacing <see cref="Location"/>'s existing file.</remarks>
        public string? TemplatePresentationId { get; set; }
        /// <summary>Gets or sets how unsupported complex slides are represented; defaults to whole-slide rasterization.</summary>
        public GoogleSlidesComplexSlideMode ComplexSlides { get; set; } = GoogleSlidesComplexSlideMode.RasterizeComplexSlides;
        /// <summary>Gets or sets revision handling for an existing presentation.</summary>
        public GoogleSlidesReplaceOptions Replace { get; set; } = new GoogleSlidesReplaceOptions();
        /// <summary>Gets or sets the policy checked against the export's fidelity report before mutation.</summary>
        public GoogleWorkspaceFidelityPolicy FidelityPolicy { get; set; } = new GoogleWorkspaceFidelityPolicy();
    }

    /// <summary>Controls how a Google presentation is loaded into OfficeIMO.</summary>
    public sealed class GoogleSlidesImportOptions {
        /// <summary>Gets or sets Drive-export or native Slides API import; defaults to Drive export.</summary>
        public GoogleWorkspaceImportMode Mode { get; set; } = GoogleWorkspaceImportMode.DriveExport;
        /// <summary>Gets or sets options used when loading the PPTX returned by Drive export.</summary>
        public PowerPointLoadOptions LoadOptions { get; set; } = new PowerPointLoadOptions();
        /// <summary>Gets or sets an optional progress observer for Drive-export transfer.</summary>
        public IProgress<OfficeIMO.GoogleWorkspace.Drive.GoogleDriveTransferProgress>? Progress { get; set; }

        /// <summary>
        /// Gets or sets the positive byte limit for each Google-hosted image during native import; defaults to 50 MiB.
        /// </summary>
        public long MaxImageBytes { get; set; } = 50L * 1024 * 1024;
    }
}
