namespace OfficeIMO.PowerPoint.GoogleSlides {
    /// <summary>Plans, prepares, and exports OfficeIMO presentations to Google Slides.</summary>
    public interface IGoogleSlidesExporter {
        /// <summary>Counts native and fallback content without contacting Google or rendering whole-slide fallback images.</summary>
        /// <remarks>Planning may still read source picture and background image bytes.</remarks>
        GoogleSlidesTranslationPlan BuildPlan(PowerPointPresentation presentation, GoogleSlidesSaveOptions? options = null);
        /// <summary>Prepares native elements and renders complex-slide images without contacting Google.</summary>
        GoogleSlidesBatch BuildBatch(PowerPointPresentation presentation, GoogleSlidesSaveOptions? options = null);
        /// <summary>Creates or replaces a Google presentation using the configured session and policy.</summary>
        Task<GooglePresentationReference> ExportAsync(PowerPointPresentation presentation, OfficeIMO.GoogleWorkspace.GoogleWorkspaceSession session, GoogleSlidesSaveOptions? options = null, CancellationToken cancellationToken = default);
    }
}
