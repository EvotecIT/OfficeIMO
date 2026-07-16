using OfficeIMO.GoogleWorkspace;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.GoogleSlides {
    public static class PowerPointGoogleSlidesExtensions {
        public static GoogleSlidesTranslationPlan BuildGoogleSlidesPlan(this PowerPointPresentation presentation, GoogleSlidesSaveOptions? options = null) => new GoogleSlidesExporter().BuildPlan(presentation, options);
        public static GoogleSlidesBatch BuildGoogleSlidesBatch(this PowerPointPresentation presentation, GoogleSlidesSaveOptions? options = null) => new GoogleSlidesExporter().BuildBatch(presentation, options);
        public static Task<GooglePresentationReference> ExportToGoogleSlidesAsync(this PowerPointPresentation presentation, GoogleWorkspaceSession session, GoogleSlidesSaveOptions? options = null, CancellationToken cancellationToken = default) =>
            new GoogleSlidesExporter().ExportAsync(presentation, session, options, cancellationToken);
        public static Task<GoogleSlidesImportResult> ImportGoogleSlidesAsync(this GoogleWorkspaceSession session, string presentationId, GoogleSlidesImportOptions? options = null, CancellationToken cancellationToken = default) =>
            new GoogleSlidesImporter().ImportAsync(presentationId, session, options, cancellationToken);
    }
}
