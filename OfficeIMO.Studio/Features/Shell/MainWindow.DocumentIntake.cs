using Avalonia.Input;
using Avalonia.Platform.Storage;
using OfficeIMO.Studio.Infrastructure.Diagnostics;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindow {
    private string[] _initialDocumentPaths = [];
    private readonly TaskCompletionSource _startupCompleted = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private Task _activationTail = Task.CompletedTask;
    private IntakeError? _intakeError;
    private sealed record IntakeError(WeakReference<MainWindowViewModel> Owner, string Message);

    internal void OpenInitialDocument(string[]? args) {
        _initialDocumentPaths = args?.Where(static argument => !string.IsNullOrWhiteSpace(argument))
            .Select(static candidate => {
                try { return System.IO.Path.GetFullPath(candidate); }
                catch (Exception) when (candidate.Length > 0) { return candidate; }
            }).ToArray() ?? [];
    }

    /// <summary>True until startup cleanup, session inspection, and any initial document have finished.</summary>
    internal bool IsStartingUp { get; private set; } = true;

    private async void OnOpened(object? sender, EventArgs e) {
        try {
            await CompleteStartupAsync();
        } catch (Exception error) when (error is not OutOfMemoryException) {
            ReportIntakeError(error);
        } finally {
            IsStartingUp = false;
            _startupCompleted.TrySetResult();
        }
    }

    private async Task CompleteStartupAsync() {
        ViewModel.SetViewportSize(PagesList.Bounds.Width, PagesList.Bounds.Height);
        try {
            var cleanup = await _services.Recovery.CleanupExpiredAsync();
            if (cleanup.FailedFiles > 0) _services.Diagnostics.Write(StudioDiagnosticLevel.Warning, "Recovery", "CleanupIncomplete");
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
            _services.Diagnostics.Write(StudioDiagnosticLevel.Warning, "Recovery", "CleanupFailed", error);
        }
        if (_windowClosed) return;
        await _session.InspectAsync();
        if (_windowClosed) return;
        foreach (string path in _initialDocumentPaths) {
            if (_windowClosed) return;
            if (!IsPdf(path) && IsConvertible(path)) {
                if (ViewModel.ConversionWorkbench.AddDroppedPaths([path])) ViewModel.ShowConversionWorkbenchCommand.Execute(null);
            } else await TabHost.OpenDocumentAsync(path);
        }
    }

    /// <summary>Accepts OS-owned references on the UI thread and transfers supported files to the window storage owner.</summary>
    internal Task OpenActivatedItemsAsync(IReadOnlyList<IStorageItem> items) {
        Avalonia.Threading.Dispatcher.UIThread.VerifyAccess();
        var files = items.Distinct<IStorageItem>(ReferenceEqualityComparer.Instance).ToArray();
        Task next = ProcessActivationAsync(files, _activationTail);
        _activationTail = next;
        return next;
    }

    private async Task ProcessActivationAsync(IStorageItem[] items, Task previous) {
        var unowned = new HashSet<IStorageItem>(items, ReferenceEqualityComparer.Instance);
        try {
            await previous;
            await _startupCompleted.Task;
            if (_windowClosed) return;
            Activate();
            IntakeError? priorError = _intakeError;
            foreach (IStorageItem item in items) {
                if (_windowClosed) return;
                if (item is not IStorageFile file || (!IsPdf(file.Name) && !IsConvertible(file.Name))) {
                    ShowIntakeError(_services.Localizer.Get("Activation.Unsupported"));
                    continue;
                }
                if (!ViewModel.CanStartDocumentTransition || (!IsPdf(file.Name) && !ViewModel.ConversionWorkbench.CanEditQueue)) {
                    ShowIntakeError(_services.Localizer.Get("Activation.Busy"));
                    continue;
                }
                try {
                    // RegisterAsync owns this item even if bookmark acquisition or shutdown fails.
                    unowned.Remove(item);
                    string location = await _services.Storage.RegisterAsync(file, CancellationToken.None);
                    if (_windowClosed) return;
                    if (IsPdf(file.Name)) {
                        await TabHost.OpenDocumentAsync(location);
                        if (ViewModel.HasDocument) ClearPriorIntakeError(priorError);
                    } else if (ViewModel.ConversionWorkbench.AddDroppedPaths([location])) {
                        ClearPriorIntakeError(priorError);
                        ViewModel.ShowConversionWorkbenchCommand.Execute(null);
                    }
                } catch (Exception error) when (error is not OutOfMemoryException) { ReportIntakeError(error); }
            }
            _services.Diagnostics.Write(StudioDiagnosticLevel.Information, "FileActivation", "BatchHandled");
        } catch (Exception error) when (error is not OutOfMemoryException) {
            ReportIntakeError(error);
        } finally {
            foreach (IStorageItem item in unowned) {
                if (_services.Storage.OwnsProviderItem(item)) continue;
                try { item.Dispose(); }
                catch (Exception error) when (error is not OutOfMemoryException) { ReportIntakeError(error); }
            }
        }
    }

    private void ReportIntakeError(Exception error) {
        _services.Diagnostics.Write(StudioDiagnosticLevel.Warning, "FileActivation", "IntakeFailed", error);
        if (!_windowClosed) ShowIntakeError(error.Message);
    }

    private void ShowIntakeError(string message) {
        ViewModel.ErrorMessage = message;
        _intakeError = new(new(ViewModel), message);
    }

    private void ClearPriorIntakeError(IntakeError? prior) {
        // Preserve errors produced by this batch and any later operation that replaced the banner.
        if (prior is null || !ReferenceEquals(_intakeError, prior)) return;
        if (prior.Owner.TryGetTarget(out var document) && document.ErrorMessage == prior.Message) document.ErrorMessage = null;
        _intakeError = null;
    }

    private void OnDragOver(object? sender, DragEventArgs e) {
        if (e.Handled) return;
        DropPlan plan = PlanDrop(e);
        e.DragEffects = plan.IsEmpty ? DragDropEffects.None : DragDropEffects.Copy;
        e.Handled = true;
        if (plan.HasFiles) ShowDropOverlay(plan);
    }

    private void OnDragLeave(object? sender, DragEventArgs e) => DropOverlay.IsVisible = false;

    private async void OnDrop(object? sender, DragEventArgs e) {
        DropOverlay.IsVisible = false;
        if (e.Handled) return;
        e.Handled = true;
        DropPlan plan = PlanDrop(e);
        if (plan.IsEmpty) return;
        try {
            if (plan.Convert.Count > 0) {
                IReadOnlyList<string> locations = await _services.Storage.RegisterManyAsync(plan.Convert, CancellationToken.None);
                if (_windowClosed) return;
                if (ViewModel.ConversionWorkbench.AddDroppedPaths(locations)) ViewModel.ShowConversionWorkbenchCommand.Execute(null);
            }
            foreach (IStorageFile file in plan.Open) {
                string location = await _services.Storage.RegisterAsync(file, CancellationToken.None);
                if (_windowClosed) return;
                await TabHost.OpenDocumentAsync(location);
            }
        } catch (Exception error) when (error is not OutOfMemoryException) {
            if (!_windowClosed) ViewModel.ErrorMessage = error.Message;
        }
    }

    private sealed record DropPlan(IReadOnlyList<IStorageFile> Open, IReadOnlyList<IStorageItem> Convert, bool HasFiles) {
        public bool IsEmpty => Open.Count == 0 && Convert.Count == 0;
    }

    // PDFs open in tabs; other supported inputs go to the conversion queue. In the conversion
    // workbench every dropped file joins the queue so PDFs can be converted too.
    private DropPlan PlanDrop(DragEventArgs e) {
        IStorageItem[] files = e.DataTransfer.TryGetFiles()?.Where(item => item is IStorageFile ||
            item is IStorageFolder && Path.GetExtension(item.Name).ToLowerInvariant() is ".pages" or ".numbers" or ".key").ToArray() ?? [];
        if (files.Length == 0) return new([], [], false);
        if (!ViewModel.CanStartDocumentTransition) return new([], [], true);
        bool canQueue = ViewModel.ConversionWorkbench.CanEditQueue;
        if (ViewModel.IsConversionMode) return new([], canQueue ? files : [], true);
        IStorageFile[] pdfs = files.OfType<IStorageFile>().Where(file => IsPdf(file.Name)).ToArray();
        IStorageItem[] others = canQueue ? files.Where(file => !IsPdf(file.Name) && IsConvertible(file.Name)).ToArray() : [];
        return new(pdfs, others, true);
    }

    private static bool IsPdf(string name) => string.Equals(System.IO.Path.GetExtension(name), ".pdf", StringComparison.OrdinalIgnoreCase);

    private bool IsConvertible(string name) => ViewModel.ConversionWorkbench.Routes.Any(route =>
        route.Route.SourceExtensions.Any(extension => string.Equals("." + extension.TrimStart('.'),
            System.IO.Path.GetExtension(name), StringComparison.OrdinalIgnoreCase)));

    private void ShowDropOverlay(DropPlan plan) {
        var text = _services.Localizer;
        (DropOverlayTitle.Text, DropOverlayDetail.Text) = plan switch {
            { IsEmpty: true } => (text.Get("Drop.Unsupported"), text.Get("Drop.UnsupportedDetail")),
            { Convert.Count: 0 } => (text.Format("Drop.OpenPdfs", plan.Open.Count), text.Get("Drop.OpenDetail")),
            { Open.Count: 0 } => (text.Format("Drop.ConvertFiles", plan.Convert.Count), text.Get("Drop.ConvertDetail")),
            _ => (text.Format("Drop.OpenAndConvert", plan.Open.Count, plan.Convert.Count), text.Get("Drop.MixedDetail"))
        };
        DropOverlay.IsVisible = true;
    }
}
