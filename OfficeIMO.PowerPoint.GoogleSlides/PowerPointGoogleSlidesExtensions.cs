using OfficeIMO.GoogleWorkspace;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.GoogleSlides {
    /// <summary>Convenience entry points for Slides planning, export, and import.</summary>
    public static class PowerPointGoogleSlidesExtensions {
        /// <summary>Reports presentation export choices without rendering whole-slide fallback images or contacting Google.</summary>
        /// <remarks>Planning may still read source picture and background image bytes.</remarks>
        public static GoogleSlidesTranslationPlan BuildGoogleSlidesPlan(this PowerPointPresentation presentation, GoogleSlidesSaveOptions? options = null) => new GoogleSlidesExporter().BuildPlan(presentation, options);
        /// <summary>Prepares an export batch, rendering complex-slide images locally when selected.</summary>
        public static GoogleSlidesBatch BuildGoogleSlidesBatch(this PowerPointPresentation presentation, GoogleSlidesSaveOptions? options = null) => new GoogleSlidesExporter().BuildBatch(presentation, options);
        /// <summary>Creates or replaces a Google presentation using the supplied session.</summary>
        public static Task<GooglePresentationReference> ExportToGoogleSlidesAsync(this PowerPointPresentation presentation, GoogleWorkspaceSession session, GoogleSlidesSaveOptions? options = null, CancellationToken cancellationToken = default) =>
            new GoogleSlidesExporter().ExportAsync(presentation, session, options, cancellationToken);
        /// <summary>Imports a Google presentation and returns a caller-owned OfficeIMO presentation.</summary>
        public static Task<GoogleSlidesImportResult> ImportGoogleSlidesAsync(this GoogleWorkspaceSession session, string presentationId, GoogleSlidesImportOptions? options = null, CancellationToken cancellationToken = default) =>
            new GoogleSlidesImporter().ImportAsync(presentationId, session, options, cancellationToken);
    }
}
