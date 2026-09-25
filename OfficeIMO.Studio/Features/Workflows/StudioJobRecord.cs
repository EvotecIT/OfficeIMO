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
    public DateTimeOffset Started { get; internal set; } = DateTimeOffset.Now;
    public string StartedLabel => Started.ToString("g");

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(CancelCommand))]
    private bool _isActive = true;
    [ObservableProperty]
    private string _status = string.Empty;
    [ObservableProperty]
    private double _progress;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(OutputName))]
    [NotifyPropertyChangedFor(nameof(CanOpenOutput))]
    private string? _outputPath;

    /// <summary>The output file or folder name; the full location stays in the tooltip.</summary>
    public string OutputName => string.IsNullOrWhiteSpace(OutputPath) ? string.Empty
        : System.IO.Path.GetFileName(OutputPath.TrimEnd('\\', '/')) is { Length: > 0 } name ? name : OutputPath;

    /// <summary>Finished outcome for the progress colour: running, succeeded, failed or cancelled.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsSucceeded), nameof(IsFailed), nameof(IsCancelled))]
    private OfficeWorkflowStatus? _outcome;

    public bool IsSucceeded => Outcome == OfficeWorkflowStatus.Completed;
    public bool IsFailed => Outcome is OfficeWorkflowStatus.Failed or OfficeWorkflowStatus.Unconfirmed;
    public bool IsCancelled => Outcome == OfficeWorkflowStatus.Cancelled;
    [ObservableProperty]
    private string? _summary;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanOpenOutput))]
    private bool _hasOutput;

    public bool CanOpenOutput => HasOutput && (!IsDirectoryOutput ||
        OutputPath is not null && OfficeIMO.Internal.OfficeStorageIdentity.GetLocalPath(OutputPath) is not null);
    private bool IsDirectoryOutput { get; set; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasRecovery))]
    private OfficeWorkflowOutputRecovery? _recovery;

    public bool HasRecovery => Recovery is not null;
    public System.Collections.ObjectModel.ObservableCollection<OfficeWorkflowOutputRecovery> Recoveries { get; } = new();
    public bool HasMultipleRecoveries => Recoveries.Count > 1;
    partial void OnRecoveryChanged(OfficeWorkflowOutputRecovery? value) => RecoveryDiscardRequested = false;

    internal void RemoveRecovery(OfficeWorkflowOutputRecovery recovery) {
        Recoveries.Remove(recovery);
        Recovery = Recoveries.FirstOrDefault();
        RecoveryDiscardRequested = false;
        OnPropertyChanged(nameof(HasMultipleRecoveries));
    }

    internal void CompleteBatch(OfficeWorkflowStatus status, string? outputPath, string summary,
        IReadOnlyList<OfficeWorkflowOutputRecovery> recoveries, bool hasVerifiedOutput, bool isDirectoryOutput) {
        if (!IsActive) return;
        IsDirectoryOutput = isDirectoryOutput;
        Complete(status, outputPath, summary, recoveries.FirstOrDefault());
        Recoveries.Clear();
        foreach (var recovery in recoveries) Recoveries.Add(recovery);
        HasOutput = hasVerifiedOutput && !string.IsNullOrWhiteSpace(outputPath);
        OnPropertyChanged(nameof(HasMultipleRecoveries));
    }

    [ObservableProperty]
    private bool _recoveryDiscardRequested;

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

    internal void Complete(OfficeWorkflowStatus status, string? outputPath, string summary, OfficeWorkflowOutputRecovery? recovery = null) {
        if (!IsActive) return;
        OutputPath = outputPath;
        HasOutput = status == OfficeWorkflowStatus.Completed && !string.IsNullOrWhiteSpace(outputPath);
        Summary = summary;
        Recovery = recovery;
        Recoveries.Clear();
        if (recovery is not null) Recoveries.Add(recovery);
        OnPropertyChanged(nameof(HasMultipleRecoveries));
        Status = status == OfficeWorkflowStatus.Unconfirmed
            ? _localizer.GetOrDefault("Jobs.Unconfirmed", "Check output")
            : _localizer.GetOrDefault("Workflow.Status." + status, status.ToString());
        Progress = status == OfficeWorkflowStatus.Cancelled ? Progress : 1D;
        Outcome = status;
        _cancel = null;
        IsActive = false;
    }

    internal void Unconfirmed(string summary) {
        if (!IsActive) return;
        Summary = summary;
        Status = _localizer.GetOrDefault("Jobs.Unconfirmed", "Check output");
        Outcome = OfficeWorkflowStatus.Unconfirmed;
        Progress = 1D;
        _cancel = null;
        IsActive = false;
    }
}
