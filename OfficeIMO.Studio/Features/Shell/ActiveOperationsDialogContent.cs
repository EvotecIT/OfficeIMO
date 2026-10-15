using OfficeIMO.Studio.Infrastructure;
using System.ComponentModel;
using System.Collections.Specialized;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Keeps the owner alive until its operations finish, with explicit cancellation or waiting.</summary>
internal sealed class ActiveOperationsDialogContent : StudioDialogContent {
    private readonly HashSet<MainWindowViewModel> _documents;
    private readonly StudioDocumentTabHost? _host;
    private readonly StudioSessionController? _session;
    private readonly TextBlock _status;
    private readonly Button _wait;
    private readonly Button _cancel;
    private readonly IStudioLocalizer _localizer;
    private readonly Workflows.StudioJobHistory? _jobs;
    private readonly StackPanel _running = new() { Spacing = 8 };
    private readonly ScrollViewer _runningList;
    private bool _closeWhenIdle;
    private bool _closed;
    private bool _cancelRequested;

    internal ActiveOperationsDialogContent(IEnumerable<MainWindowViewModel> documents, IStudioLocalizer localizer,
        StudioDocumentTabHost? host = null, StudioSessionController? session = null, Workflows.StudioJobHistory? jobs = null) {
        _documents = documents.ToHashSet();
        _jobs = jobs;
        _runningList = new ScrollViewer { Name = "RunningWork", Content = _running, MaxHeight = 220, VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto };
        _host = host;
        _session = session;
        _localizer = localizer;
        Title = Text("Title", "Work is still running");
        Width = 540;
        SizeToContent = SizeToContent.Height;
        CanResize = false;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        _status = new TextBlock { Text = Text("Description", "Choose whether to wait for the work to finish or request cancellation before closing."), TextWrapping = TextWrapping.Wrap };
        _wait = new Button { Name = "WaitAndClose", Content = Text("Wait", "Wait and close"), Classes = { "primary" } };
        _cancel = new Button { Name = "CancelWorkAndClose", Content = Text("Cancel", "Cancel work and close"), Classes = { "tool" } };
        var keep = new Button { Name = "KeepOpen", Content = Text("Keep", "Keep open"), Classes = { "tool" }, IsCancel = true };
        keep.Click += (_, _) => Close(false);
        _wait.Click += (_, _) => BeginClose(cancel: false);
        _cancel.Click += (_, _) => BeginClose(cancel: true);
        Content = new StackPanel {
            Margin = new Thickness(24), Spacing = 16,
            Children = {
                new TextBlock { Text = Title, FontSize = 20, FontWeight = FontWeight.SemiBold, TextWrapping = TextWrapping.Wrap },
                _status,
                _runningList,
                new TextBlock { Text = Text("Outputs", "Outputs already saved are kept. Cancellation does not undo them. Job history lasts only for this session and work will not resume after restarting."), TextWrapping = TextWrapping.Wrap },
                new WrapPanel { Orientation = Orientation.Horizontal, Children = { keep, _wait, _cancel } }
            }
        };
        foreach (Button button in new[] { keep, _wait, _cancel }) button.Margin = new Thickness(0, 4, 8, 4);
        foreach (MainWindowViewModel document in _documents) document.PropertyChanged += OnDocumentChanged;
        if (_host is not null) _host.Tabs.CollectionChanged += OnTabsChanged;
        if (_session is not null) _session.PropertyChanged += OnSessionChanged;
        if (_jobs is not null) {
            foreach (Workflows.StudioJobRecord record in _jobs.Entries) record.PropertyChanged += OnJobChanged;
            ((INotifyCollectionChanged)_jobs.Entries).CollectionChanged += OnJobsChanged;
        }
        ShowRunningWork();
        Opened += (_, _) => keep.Focus();
        Closed += (_, _) => {
            _closed = true;
            foreach (MainWindowViewModel document in _documents) document.PropertyChanged -= OnDocumentChanged;
            if (_host is not null) _host.Tabs.CollectionChanged -= OnTabsChanged;
            if (_session is not null) _session.PropertyChanged -= OnSessionChanged;
            if (_jobs is not null) {
                foreach (Workflows.StudioJobRecord record in _jobs.Entries) record.PropertyChanged -= OnJobChanged;
                ((INotifyCollectionChanged)_jobs.Entries).CollectionChanged -= OnJobsChanged;
            }
        };
    }

    // The decision is about specific work: list what is running so waiting or cancelling is an informed choice.
    private void ShowRunningWork() {
        _running.Children.Clear();
        foreach (MainWindowViewModel document in _documents.Where(document => document.CanCancelOperation))
            _running.Children.Add(Row(document.DocumentName.TrimEnd(' ', '*'), document.OperationStatus ?? Text("Working", "Working…"), null));
        foreach (Workflows.StudioJobRecord record in _jobs?.Entries.Where(record => record.IsActive) ?? [])
            _running.Children.Add(Row(record.Title + " · " + System.IO.Path.GetFileName(record.InputLabel), record.Status, record.Progress));
        _runningList.IsVisible = _running.Children.Count > 0;
    }

    private static Control Row(string title, string status, double? progress) {
        var row = new StackPanel { Spacing = 2 };
        row.Children.Add(new TextBlock { Text = title, FontWeight = FontWeight.SemiBold, TextTrimming = TextTrimming.CharacterEllipsis });
        row.Children.Add(new TextBlock { Text = status, TextWrapping = TextWrapping.Wrap, Classes = { "muted" } });
        if (progress is { } value) row.Children.Add(new ProgressBar { Minimum = 0, Maximum = 1, Value = value, Height = 6, MinWidth = 0 });
        return row;
    }

    private void OnJobChanged(object? sender, PropertyChangedEventArgs args) {
        if (args.PropertyName is nameof(Workflows.StudioJobRecord.IsActive) or nameof(Workflows.StudioJobRecord.Status) or nameof(Workflows.StudioJobRecord.Progress))
            this.Dispatcher.Post(CheckIdle);
    }

    // Work started while the dialog is open is listed, and the window waits for it too.
    private void OnJobsChanged(object? sender, NotifyCollectionChangedEventArgs args) {
        foreach (Workflows.StudioJobRecord record in args.OldItems?.OfType<Workflows.StudioJobRecord>() ?? []) record.PropertyChanged -= OnJobChanged;
        foreach (Workflows.StudioJobRecord record in args.NewItems?.OfType<Workflows.StudioJobRecord>() ?? []) record.PropertyChanged += OnJobChanged;
        this.Dispatcher.Post(CheckIdle);
    }

    private string Text(string key, string fallback) => _localizer.GetOrDefault("Dialog.ActiveOperations." + key, fallback);

    private void BeginClose(bool cancel) {
        _closeWhenIdle = true;
        _wait.IsEnabled = false;
        _status.Text = cancel
            ? Text("Cancelling", "Cancellation requested. Waiting for operations to stop and finish output cleanup…")
            : Text("Waiting", "Waiting for operations to finish. You can still request cancellation or keep this window open.");
        if (cancel) {
            _cancelRequested = true;
            _cancel.IsEnabled = false;
            _session?.CancelActiveOperation();
            foreach (MainWindowViewModel document in _documents.ToArray()) document.CancelCurrentOperation();
            CancelJobs();
        }
        CheckIdle();
    }

    private void OnTabsChanged(object? sender, NotifyCollectionChangedEventArgs args) {
        if (_host is not null) {
            foreach (MainWindowViewModel document in _host.OperationDocuments) {
                if (_documents.Add(document)) document.PropertyChanged += OnDocumentChanged;
            }
        }
        this.Dispatcher.Post(CheckIdle);
    }

    private void OnSessionChanged(object? sender, PropertyChangedEventArgs args) {
        if (args.PropertyName == nameof(StudioSessionController.IsBusy)) this.Dispatcher.Post(CheckIdle);
    }

    private void OnDocumentChanged(object? sender, PropertyChangedEventArgs args) {
        if (args.PropertyName == nameof(MainWindowViewModel.CanCancelOperation)) {
            // Finish after the operation's finally block and its remaining UI notifications unwind.
            this.Dispatcher.Post(CheckIdle);
        }
    }

    private void CheckIdle() {
        if (_closed) return;
        ShowRunningWork();
        if (!_closeWhenIdle) return;
        if (_cancelRequested) {
            _session?.CancelActiveOperation();
            foreach (MainWindowViewModel document in _documents.Where(document => document.CanCancelOperation).ToArray()) document.CancelCurrentOperation();
            CancelJobs();
        }
        if (_session?.IsBusy != true && _documents.All(document => !document.CanCancelOperation) && _jobs?.HasActive != true) Close(true);
    }

    private void CancelJobs() {
        foreach (Workflows.StudioJobRecord record in _jobs?.Entries.Where(record => record.IsActive).ToArray() ?? [])
            if (record.CancelCommand.CanExecute(null)) record.CancelCommand.Execute(null);
    }
}
