namespace OfficeIMO.PowerPoint.GoogleSlides {
    public interface IGoogleSlidesImporter {
        Task<GoogleSlidesImportResult> ImportAsync(string presentationId, OfficeIMO.GoogleWorkspace.GoogleWorkspaceSession session, GoogleSlidesImportOptions? options = null, CancellationToken cancellationToken = default);
    }
}
