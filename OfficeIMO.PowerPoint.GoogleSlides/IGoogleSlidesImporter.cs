namespace OfficeIMO.PowerPoint.GoogleSlides {
    /// <summary>Imports a Google presentation into a caller-owned OfficeIMO presentation.</summary>
    public interface IGoogleSlidesImporter {
        /// <summary>Loads a presentation through Drive export or native Slides projection.</summary>
        Task<GoogleSlidesImportResult> ImportAsync(string presentationId, OfficeIMO.GoogleWorkspace.GoogleWorkspaceSession session, GoogleSlidesImportOptions? options = null, CancellationToken cancellationToken = default);
    }
}
