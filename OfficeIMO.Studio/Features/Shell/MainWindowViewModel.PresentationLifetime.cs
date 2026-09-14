namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private bool _presentationActive = true;

    /// <summary>Releases recreatable page data when a tab is hidden; document bytes and undo remain intact.</summary>
    internal void SetPresentationActive(bool active) {
        _presentationActive = active;
        if (!active) {
            foreach (var page in Pages) page.DetachFromViewport();
            foreach (var page in OrganizerPages) page.Detach();
            foreach (var page in ComparisonPages) page.DetachFromViewport();
        }
        ApplyPresentationCachePolicy();
    }

    private void ApplyPresentationCachePolicy() {
        _sceneCoordinator?.SetCacheEnabled(_presentationActive);
        _renderCoordinator?.SetCacheEnabled(_presentationActive);
        _comparisonSceneCoordinator?.SetCacheEnabled(_presentationActive);
        _comparisonRenderCoordinator?.SetCacheEnabled(_presentationActive);
    }
}
