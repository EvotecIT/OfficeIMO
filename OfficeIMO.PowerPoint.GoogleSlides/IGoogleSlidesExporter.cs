namespace OfficeIMO.PowerPoint.GoogleSlides {
    /// <summary>Plans, prepares, and exports OfficeIMO presentations to Google Slides.</summary>
    public interface IGoogleSlidesExporter {
        /// <summary>Counts native and fallback content without contacting Google or materializing raster images.</summary>
        GoogleSlidesTranslationPlan BuildPlan(PowerPointPresentation presentation, GoogleSlidesSaveOptions? options = null);
        /// <summary>Prepares native elements and renders complex-slide images without contacting Google.</summary>
        GoogleSlidesBatch BuildBatch(PowerPointPresentation presentation, GoogleSlidesSaveOptions? options = null);
        /// <summary>Creates or replaces a Google presentation using the configured session and policy.</summary>
        Task<GooglePresentationReference> ExportAsync(PowerPointPresentation presentation, OfficeIMO.GoogleWorkspace.GoogleWorkspaceSession session, GoogleSlidesSaveOptions? options = null, CancellationToken cancellationToken = default);
    }
}
