using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>An in-memory presentation record for one workflow attempt.</summary>
public sealed partial class StudioJobRecord : ObservableObject {
    private readonly IStudioLocalizer _localizer;
    private Action? _cancel;
    internal StudioJobRecord(string title, string input, string? destination, Action cancel, bool batch, IStudioLocalizer localizer) {
        Title = title;
        Input = input;
        Destination = destination;
        _cancel = cancel;
        _localizer = localizer;
        CancelLabel = localizer.GetOrDefault(batch ? "Jobs.CancelBatch" : "Jobs.Cancel", batch ? "Cancel batch" : "Cancel job");
        Status = localizer.GetOrDefault("Jobs.Queued", "Queued");
    }

    public string Title { get; }
    public string Input { get; }
    public string InputLabel => Input.Replace('\r', ' ').Replace('\n', ' ');
    public string? Destination { get; }
    public string CancelLabel { get; }
    public DateTimeOffset Started { get; } = DateTimeOffset.Now;
    public string StartedLabel => Started.ToString("g");

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(CancelCommand))]
    private bool _isActive = true;
    [ObservableProperty]
    private string _status = string.Empty;
    [ObservableProperty]
    private double _progress;
    [ObservableProperty]
    private string? _outputPath;
    [ObservableProperty]
    private string? _summary;
    [ObservableProperty]
    private bool _hasOutput;

    [RelayCommand(CanExecute = nameof(IsActive))]
    private void Cancel() {
        _cancel?.Invoke();
        if (IsActive) Status = _localizer.GetOrDefault("Jobs.Cancelling", "Cancellation requested");
    }

    internal void Report(OfficeWorkflowProgress progress) {
        if (!IsActive) return;
        Progress = progress.Fraction;
        Status = _localizer.GetOrDefault("Jobs.Running", "Running");
        Summary = progress.Message;
    }

    internal void Complete(OfficeWorkflowStatus status, string? outputPath, string summary) {
        if (!IsActive) return;
        OutputPath = outputPath;
        HasOutput = status == OfficeWorkflowStatus.Completed && !string.IsNullOrWhiteSpace(outputPath);
        Summary = summary;
        Status = _localizer.GetOrDefault("Workflow.Status." + status, status.ToString());
        Progress = status == OfficeWorkflowStatus.Cancelled ? Progress : 1D;
        _cancel = null;
        IsActive = false;
    }

    internal void Unconfirmed(string summary) {
        if (!IsActive) return;
        Summary = summary;
        Status = _localizer.GetOrDefault("Jobs.Unconfirmed", "Check output");
        _cancel = null;
        IsActive = false;
    }
}
