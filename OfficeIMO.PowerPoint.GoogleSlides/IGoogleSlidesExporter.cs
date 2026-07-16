namespace OfficeIMO.PowerPoint.GoogleSlides {
    public interface IGoogleSlidesExporter {
        GoogleSlidesTranslationPlan BuildPlan(PowerPointPresentation presentation, GoogleSlidesSaveOptions? options = null);
        GoogleSlidesBatch BuildBatch(PowerPointPresentation presentation, GoogleSlidesSaveOptions? options = null);
        Task<GooglePresentationReference> ExportAsync(PowerPointPresentation presentation, OfficeIMO.GoogleWorkspace.GoogleWorkspaceSession session, GoogleSlidesSaveOptions? options = null, CancellationToken cancellationToken = default);
    }
}
