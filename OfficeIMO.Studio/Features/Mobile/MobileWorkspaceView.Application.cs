using Avalonia;
using Avalonia.Controls;
using Avalonia.Data;
using Avalonia.Interactivity;
using Avalonia.Threading;
using OfficeIMO.Studio.Features.Home;
using OfficeIMO.Studio.Features.Settings;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Features.Mobile;

public sealed partial class MobileWorkspaceView {
    private readonly Dictionary<StudioWorkspaceMode, Control> _features = [];
    private bool? _wideApplicationNavigation;
    private DocumentWorkspaceView? _editingWorkspace;
    private bool IsTouchReader => Document is null || Document.IsPdfWorkspaceMode && Document.IsViewDocumentMode;

    private void OnCloseAssistantClick(object? sender, RoutedEventArgs e) {
        if (Document is { } document) document.IsAssistantVisible = false;
    }

    private void OnApplicationMenuClick(object? sender, RoutedEventArgs e) =>
        ApplicationNavigation.IsPaneOpen = !ApplicationNavigation.IsPaneOpen;

    private void OnApplicationNavigationClick(object? sender, RoutedEventArgs e) {
        if (ApplicationNavigation.DisplayMode == SplitViewDisplayMode.Overlay) ApplicationNavigation.IsPaneOpen = false;
    }

    private void UpdateApplicationFeature() {
        ApplyKeyboardAvoidance();
        UpdateDialogLayout();
        bool wide = Bounds.Width >= 1000;
        AssistantHost.OpenPaneLength = Math.Max(1, Math.Min(380, Bounds.Width - 16));
        if (_wideApplicationNavigation != wide) {
            _wideApplicationNavigation = wide;
            ApplicationNavigation.DisplayMode = wide ? SplitViewDisplayMode.Inline : SplitViewDisplayMode.Overlay;
            ApplicationNavigation.IsPaneOpen = wide;
        }
        bool reader = IsTouchReader;
        UpdateDocumentStatus();
        bool documentVisible = Document?.IsPdfWorkspaceMode == true && Document.HasDocument;
        SaveButton.IsVisible = documentVisible;
        ShareButton.IsVisible = documentVisible && _shareDocumentAsync is not null;
        DocumentToolsButton.IsVisible = documentVisible && reader;
        HomeSampleButton.IsVisible = Document?.IsHomeMode == true && _openSampleAsync is not null;
        ReaderSurface.IsVisible = reader;
        WelcomeSurface.IsVisible = reader && Document?.IsEmpty == true;
        FeatureSurface.IsVisible = !reader;
        ReaderNavigation.IsVisible = reader && Document?.HasDocument == true;
        PagesButton.IsVisible = reader;
        OpenButton.IsVisible = Bounds.Width >= 600;
        if (!reader && Document is { } document) {
            if (!_features.TryGetValue(document.WorkspaceMode, out Control? feature)) {
                feature = CreateFeature(document.WorkspaceMode);
                _features.Add(document.WorkspaceMode, feature);
            }
            FeatureSurface.Content = feature;
        } else FeatureSurface.Content = null;
        if (!reader) SearchPanel.IsVisible = false;
        Dispatcher.UIThread.Post(UpdateActiveViewport, DispatcherPriority.Loaded);
    }

    private Control CreateFeature(StudioWorkspaceMode mode) => mode switch {
        StudioWorkspaceMode.Home => CreateHome(),
        StudioWorkspaceMode.Tools => new ToolsView(),
        StudioWorkspaceMode.Convert => new ConversionWorkbenchView(),
        StudioWorkspaceMode.Jobs => BindFeature(new StudioJobsView(), nameof(MainWindowViewModel.Jobs)),
        StudioWorkspaceMode.Settings => CreateSettings(),
        StudioWorkspaceMode.Output => new OutputIntakeWorkbenchView(),
        StudioWorkspaceMode.DocumentHealth => new DocumentHealthView(),
        StudioWorkspaceMode.Provenance => new ProvenanceWorkbenchView(),
        StudioWorkspaceMode.Invoices => BindFeature(new InvoiceWorkbenchView(), nameof(MainWindowViewModel.InvoiceWorkbench)),
        StudioWorkspaceMode.Publishing => BindFeature(new BookWorkbenchView(), nameof(MainWindowViewModel.BookWorkbench)),
        StudioWorkspaceMode.Ocr => new TabControl {
            Classes = { "pageTabs" }, Items = {
                new TabItem { Header = Infrastructure.Localization.StudioLocalization.Current.Get("OcrSession.SinglePdf"), Content = new SearchablePdfOcrView() },
                new TabItem { Header = Infrastructure.Localization.StudioLocalization.Current.Get("OcrSession.FilesTab"), Content = BindFeature(new OcrSessionView(), nameof(MainWindowViewModel.OcrSession)) }
            }
        },
        StudioWorkspaceMode.PdfWorkspace => CreateEditingWorkspace(),
        _ => throw new ArgumentOutOfRangeException(nameof(mode))
    };

    private static HomeView CreateHome() {
        var view = new HomeView();
        view.UseTouchPresentation();
        return view;
    }

    private static Control CreateSettings() {
        var view = new SettingsView();
        view.UseTouchPresentation();
        return BindFeature(view, nameof(MainWindowViewModel.Settings));
    }

    private static Control BindFeature(Control view, string property) {
        var container = new Border { Child = view };
        view.Bind(DataContextProperty, new Binding(property));
        return container;
    }

    private DocumentWorkspaceView CreateEditingWorkspace() {
        var view = _editingWorkspace = new DocumentWorkspaceView { UseTouchPresentation = true };
        view.FindControl<Button>("SaveButton")!.IsVisible = false;
        view.PagesListControl.SizeChanged += (_, _) => UpdateActiveViewport();
        view.PagesListControl.SelectionChanged += (_, _) => {
            if (!view.IsEffectivelyVisible || Document is not { } document ||
                view.PagesListControl.SelectedItem is not Reader.PdfPageViewModel page) return;
            document.SelectedPage = page;
            view.PagesListControl.ScrollIntoView(page);
        };
        view.OrganizerListControl.SelectionChanged += (_, e) => {
            if (view.IsEffectivelyVisible) Document?.UpdateOrganizerSelection(
                e.AddedItems.OfType<Organizer.PdfOrganizerPageViewModel>(),
                e.RemovedItems.OfType<Organizer.PdfOrganizerPageViewModel>());
        };
        return view;
    }

    private void UpdateActiveViewport() {
        if (Document is not { } document) return;
        Control viewport = IsTouchReader ? PageScroll : (Control?)_editingWorkspace?.PagesListControl ?? PageScroll;
        if (viewport.IsEffectivelyVisible && viewport.Bounds.Width > 0 && viewport.Bounds.Height > 0)
            document.SetViewportSize(viewport.Bounds.Width, viewport.Bounds.Height);
    }
}
