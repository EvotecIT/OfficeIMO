using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    // A pane has view state only. The tab still owns the workspace, source bytes, journal and publication guard.
    internal ViewerZoomMode ReaderZoomMode => _zoomMode;
    internal Infrastructure.Localization.IStudioLocalizer PaneLocalizer => _localizer;
    internal bool IsReplacingReaderPresentation { get; private set; }
    internal event EventHandler? ReaderPresentationReplaced;

    internal IReadOnlyList<PdfPageViewModel> CreatePanePages(Action<int> activatePage) {
        if (_session is null || _sceneCoordinator is null || _renderCoordinator is null) return [];
        return _session.Pages.Select(page => {
            var presentation = new PdfPageViewModel(page.PageNumber, page.Width, page.Height, page.RotationDegrees,
                Zoom, _sceneCoordinator, _renderCoordinator, _localizer, page.Geometry);
            presentation.LinkActivated += target => { activatePage(page.PageNumber); OnPageLinkActivated(target); };
            presentation.EditorGestureCompleted += gesture => { activatePage(page.PageNumber); OnPageEditorGestureCompleted(gesture); };
            presentation.MarkupRequested += (tool, gesture) => { activatePage(page.PageNumber); OnPageMarkupRequested(tool, gesture); };
            presentation.InlineFormNavigationRequested += direction => { activatePage(page.PageNumber); OnInlineFormNavigationRequested(direction); };
            presentation.ObjectSelected += selection => { activatePage(page.PageNumber); OnPageObjectSelected(selection); };
            presentation.AnnotationSelectionRequested += request => { activatePage(page.PageNumber); OnPageAnnotationsSelected(request); };
            presentation.AnnotationKeyRequested += (key, modifiers) => { activatePage(page.PageNumber); OnAnnotationKeyRequested(key, modifiers); };
            presentation.ObjectTransformCompleted += gesture => { activatePage(page.PageNumber); OnPageObjectTransform(gesture); };
            return presentation;
        }).ToArray();
    }

    internal void ApplyPaneNavigation(int page, double zoom, ReaderLayoutMode layout, double width, double height) {
        SelectedReaderLayoutChoice = ReaderLayoutChoices.Single(choice => choice.Mode == layout);
        if (Pages.Count > 0) SelectedPage = Pages[Math.Clamp(page - 1, 0, Pages.Count - 1)];
        SetViewportSize(width, height);
        SetTouchZoom(zoom);
    }
}
