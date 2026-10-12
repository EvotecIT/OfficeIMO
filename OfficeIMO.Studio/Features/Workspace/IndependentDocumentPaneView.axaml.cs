using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Features.Workspace;

public sealed partial class IndependentDocumentPaneView : UserControl {
    private readonly HashSet<PdfPageCanvas> _marqueePreviews = [];
    public IndependentDocumentPaneView() {
        InitializeComponent();
        AddHandler(PointerPressedEvent, OnPanePointerPressed, RoutingStrategies.Tunnel, handledEventsToo: true);
        AddHandler(PdfPageCanvas.AnnotationMarqueeEvent, OnAnnotationMarquee);
        DetachedFromVisualTree += (_, _) => ClearMarquee();
        DataContextChanged += (_, _) => ClearMarquee();
        AddHandler(GotFocusEvent, (_, _) => Model?.Activate(), RoutingStrategies.Tunnel);
    }
    private void ClearMarquee() {
        foreach (var canvas in _marqueePreviews) canvas.AnnotationMarqueePreview = null;
        _marqueePreviews.Clear();
    }
    private void OnAnnotationMarquee(object? sender, PdfAnnotationMarqueeEventArgs args) {
        ClearMarquee();
        if (args.Cancelled || Model is not { } model) return;
        var reader = args.Canvas.FindAncestorOfType<ListBox>();
        var viewport = args.Canvas.FindAncestorOfType<ScrollContentPresenter>();
        if (reader is null || viewport is null || (reader != PanePages && reader != PaneGridPages)) return;
        var selections = PdfAnnotationMarqueeProjection.Project(reader, viewport, args, _marqueePreviews);
        if (selections is null) return;
        if (args.Completed) {
            model.Activate();
            model.Document.OnPageAnnotationsSelected(new(selections, args.Additive, Toggle: false));
        }
        args.Handled = true;
    }
    private StudioDocumentPaneViewModel? Model => DataContext as StudioDocumentPaneViewModel;
    internal void FocusReader() => (Model?.IsGrid == true ? PaneGridPages : PanePages).Focus(NavigationMethod.Directional);
    private void OnPanePointerPressed(object? sender, PointerPressedEventArgs args) {
        if (args.Source is Control source && (source as PdfPageView ?? source.FindAncestorOfType<PdfPageView>())?.DataContext is PdfPageViewModel page)
            Model?.ActivatePage(page.PageNumber);
        else Model?.Activate();
    }
    private void OnViewportSizeChanged(object? sender, SizeChangedEventArgs args) => Model?.SetViewport(args.NewSize.Width, args.NewSize.Height);
    private void OnPageSelectionChanged(object? sender, SelectionChangedEventArgs args) {
        if (ReferenceEquals(args.Source, PanePages) && PanePages.SelectedItem is PdfPageViewModel page && Model is { } model) {
            if (!ReferenceEquals(model.SelectedPage, page)) model.ActivatePage(page.PageNumber);
            PanePages.ScrollIntoView(page);
        } else if (ReferenceEquals(args.Source, PaneGridPages) && PaneGridPages.SelectedItem is ReaderGridRowViewModel row) {
            PaneGridPages.ScrollIntoView(row);
        }
    }

    private void OnPageNumberKeyDown(object? sender, KeyEventArgs args) {
        if (args.Key == Key.Enter) {
            if (Model?.NavigateTypedPage(PanePageNumber.Text) == true) FocusReader();
            args.Handled = true;
        } else if (args.Key == Key.Escape && Model is { } model) {
            PanePageNumber.Text = model.SelectedPage?.PageNumber.ToString(); FocusReader(); args.Handled = true;
        }
    }
}
