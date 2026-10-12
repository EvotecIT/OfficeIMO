using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Associates two viewports with existing tabs and routes focus to the tab's canonical command owner.</summary>
public sealed partial class StudioDocumentPanes : ObservableObject, IDisposable {
    private readonly StudioDocumentTabHost _tabs;
    private bool _selecting;
    private bool _disposed;
    private StudioDocumentTabViewModel? _openingTab;
    internal StudioDocumentPanes(StudioDocumentTabHost tabs) {
        _tabs = tabs;
        tabs.PropertyChanged += OnTabHostChanged;
        tabs.Tabs.CollectionChanged += OnTabsChanged;
    }
    public ObservableCollection<StudioDocumentTabViewModel> Tabs => _tabs.Tabs;
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(IsSplit))] private StudioDocumentPaneViewModel? _left;
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(IsSplit))] private StudioDocumentPaneViewModel? _right;
    [ObservableProperty] private StudioDocumentPaneViewModel? _activePane;
    public bool IsSplit => Left is not null && Right is not null;
    public string SplitLabel => _tabs.ActiveDocument.PaneLocalizer.GetOrDefault("Panes.Split", "Independent panes");
    public string HelpLabel => _tabs.ActiveDocument.PaneLocalizer.GetOrDefault("Panes.Help", "Each pane keeps its page, zoom and layout. Editing commands apply to the active pane.");
    public string SwitchLabel => _tabs.ActiveDocument.PaneLocalizer.GetOrDefault("Panes.Switch", "Switch active pane (F6)");

    [RelayCommand] private void Split() => OpenSecondPane(_tabs.SelectedTab);
    internal void OpenSecondPane(StudioDocumentTabViewModel? tab) {
        if (_disposed || tab is null || !Tabs.Contains(tab) || !tab.Document.HasDocument) return;
        var leftTab = Left?.Tab ?? _tabs.SelectedTab ?? tab;
        // Comparison temporarily changes the reader layout; restore it before a new pane captures view state.
        leftTab.Document.CloseComparisonCommand.Execute(null);
        tab.Document.CloseComparisonCommand.Execute(null);
        if (Left is null) Left = new(this, leftTab);
        Right?.Dispose();
        Right = new(this, tab);
        UpdatePresentedDocuments();
        Activate(Right);
    }
    internal void AssignDocument(StudioDocumentPaneViewModel pane, StudioDocumentTabViewModel tab) {
        if (_disposed || (pane != Left && pane != Right) || !Tabs.Contains(tab)) return;
        if (!tab.Document.HasDocument) {
            Activate(pane);
            _tabs.SelectedTab = tab;
            ObserveOpeningTab(tab);
            return;
        }
        tab.Document.CloseComparisonCommand.Execute(null);
        var replacement = new StudioDocumentPaneViewModel(this, tab);
        if (pane == Left) Left = replacement;
        else if (pane == Right) Right = replacement;
        else { replacement.Dispose(); return; }
        pane.Dispose();
        UpdatePresentedDocuments();
        Activate(replacement);
    }
    internal void Activate(StudioDocumentPaneViewModel pane) {
        if (_disposed || (pane != Left && pane != Right)) return;
        if (ReferenceEquals(ActivePane, pane)) return;
        if (ActivePane is { } previous) previous.IsActive = false;
        ActivePane = pane;
        _selecting = true;
        try {
            _tabs.SelectedTab = pane.Tab;
            pane.Document.WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;
            // Rebinding the canonical reader can publish its former page; keep the pane's view state until rebinding finishes.
            pane.IsActive = true;
            pane.PublishNavigation();
        } finally { _selecting = false; }
        UpdatePresentedDocuments();
    }
    internal void SwitchPane() {
        if (IsSplit) Activate(ActivePane == Left ? Right! : Left!);
    }
    internal void ClosePane(StudioDocumentPaneViewModel pane) {
        if (pane != Left && pane != Right) return;
        var survivor = pane == Left ? Right : Left;
        pane.Dispose();
        Left = null; Right = null; ActivePane = null;
        if (survivor is not null) {
            _selecting = true;
            try { _tabs.SelectedTab = survivor.Tab; survivor.IsActive = true; survivor.PublishNavigation(); }
            finally { _selecting = false; }
            survivor.Dispose();
        }
        UpdatePresentedDocuments();
    }
    private void OnTabHostChanged(object? sender, PropertyChangedEventArgs args) {
        if (!IsSplit || _selecting || args.PropertyName != nameof(StudioDocumentTabHost.SelectedTab) || _tabs.SelectedTab is not { } selected) return;
        StopObservingOpeningTab();
        if (!selected.Document.HasDocument) {
            // The tab core selects its candidate before loading it. Keep the current pane if loading fails or is cancelled.
            ObserveOpeningTab(selected);
            return;
        }
        if (ActivePane?.Tab == selected) return;
        if (Left?.Tab == selected) Activate(Left);
        else if (Right?.Tab == selected) Activate(Right);
        else if (ActivePane is { } active) AssignDocument(active, selected);
    }
    private void OnTabsChanged(object? sender, NotifyCollectionChangedEventArgs args) {
        if (_openingTab is { } opening && !Tabs.Contains(opening)) StopObservingOpeningTab();
        if (Left is { } left && !Tabs.Contains(left.Tab)) ClosePane(left);
        else if (Right is { } right && !Tabs.Contains(right.Tab)) ClosePane(right);
    }
    private void OnOpeningPresentationReplaced(object? sender, EventArgs args) {
        if (_openingTab is not { } opened || !opened.Document.HasDocument) return;
        StopObservingOpeningTab();
        if (IsSplit && _tabs.SelectedTab == opened && ActivePane is { } active) AssignDocument(active, opened);
    }
    private void ObserveOpeningTab(StudioDocumentTabViewModel tab) {
        StopObservingOpeningTab();
        _openingTab = tab;
        tab.Document.ReaderPresentationReplaced += OnOpeningPresentationReplaced;
    }
    private void StopObservingOpeningTab() {
        if (_openingTab is null) return;
        _openingTab.Document.ReaderPresentationReplaced -= OnOpeningPresentationReplaced;
        _openingTab = null;
    }
    private void UpdatePresentedDocuments() => _tabs.Core.SetPresentedDocuments(IsSplit
        ? new[] { Left!.Document, Right!.Document }.Distinct() : []);
    public void Dispose() {
        if (_disposed) return; _disposed = true;
        StopObservingOpeningTab();
        _tabs.PropertyChanged -= OnTabHostChanged;
        _tabs.Tabs.CollectionChanged -= OnTabsChanged;
        Left?.Dispose(); Right?.Dispose(); Left = null; Right = null; ActivePane = null;
        UpdatePresentedDocuments();
    }
}
