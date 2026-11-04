using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Workflows;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>Presentation state of an in-memory conversion attempt.</summary>
public enum ConversionJobState {
    /// <summary>Waiting for its first execution or an explicitly requested retry.</summary>
    Queued,
    /// <summary>The owner has reported execution progress without a terminal result.</summary>
    Running,
    /// <summary>The owner confirmed completion and any output publication.</summary>
    Completed,
    /// <summary>The owner confirmed failure without a committed output.</summary>
    Failed,
    /// <summary>The owner confirmed cancellation or left this request unstarted.</summary>
    Cancelled,
    /// <summary>No trustworthy terminal result was received; publication must be checked before repeating.</summary>
    Unconfirmed
}

public sealed partial class ConversionJobViewModel : ObservableObject {
    private readonly IStudioLocalizer _localizer;

    public ConversionJobViewModel(string inputPath, ConversionRouteChoice route) : this(inputPath, route, null) { }

    internal ConversionJobViewModel(string inputPath, ConversionRouteChoice route, IStudioLocalizer? localizer, string? fileName = null) {
        Id = Guid.NewGuid().ToString("N");
        InputPath = inputPath;
        FileName = fileName ?? Path.GetFileName(inputPath);
        Route = route;
        _localizer = localizer ?? StudioLocalization.Current;
        Status = _localizer.GetOrDefault("Conversion.Job.Queued", "Queued");
    }

    public string Id { get; }
    public string InputPath { get; }
    public string FileName { get; }
    public ConversionRouteChoice Route { get; }
    public string RouteLabel => Route.Route.Source + " → " + Route.Route.Target;
    public string Engine => Route.Engine;
    public string Fidelity => UsesRasterPages || SupportsIWorkOptions && IWorkMode == OfficeIMO.IWork.IWorkConversionMode.VisualOnly ? T("VisualFidelity", "Visual page images · review rendering and editability limits") : Route.Fidelity;
    public string KnownLimitations => Route.KnownLimitations;

    public OfficeConversionProfile OutputProfile { get; internal set; } = OfficeConversionProfile.Faithful;

    [ObservableProperty]
    private string _status = string.Empty;

    [ObservableProperty]
    private double _progressFraction;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasOutput))]
    private string? _outputPath;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasSummary))]
    private string? _summary;

    public bool HasSummary => !string.IsNullOrWhiteSpace(Summary);

    [ObservableProperty]
    private IReadOnlyList<OfficeWorkflowDiagnostic> _diagnostics = Array.Empty<OfficeWorkflowDiagnostic>();

    public bool HasWarnings => Diagnostics.Any(item => item.Severity == OfficeWorkflowDiagnosticSeverity.Warning);

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasOutput))]
    [NotifyPropertyChangedFor(nameof(HasStarted))]
    private ConversionJobState _state;


    /// <summary>Progress is shown once a job has started; a waiting job only shows its status.</summary>
    public bool HasStarted => State != ConversionJobState.Queued;

    public bool HasOutput => State == ConversionJobState.Completed && !string.IsNullOrWhiteSpace(OutputPath);

    /// <summary>The finished state shared with OCR and Jobs; warnings and fidelity loss make a completion one to review.</summary>
    internal Infrastructure.StudioOutcome Outcome => State switch {
        ConversionJobState.Completed => HasWarnings || HasLoss ? Infrastructure.StudioOutcome.CompletedWithNotes : Infrastructure.StudioOutcome.Succeeded,
        ConversionJobState.Failed => Infrastructure.StudioOutcome.Failed,
        ConversionJobState.Cancelled => Infrastructure.StudioOutcome.Cancelled,
        ConversionJobState.Unconfirmed => Infrastructure.StudioOutcome.NeedsChecking,
        _ => Infrastructure.StudioOutcome.None
    };
    public bool IsSucceeded => Outcome == Infrastructure.StudioOutcome.Succeeded;
    public bool HasNotes => Outcome == Infrastructure.StudioOutcome.CompletedWithNotes;
    public bool IsFailed => Outcome == Infrastructure.StudioOutcome.Failed;
    public bool IsCancelled => Outcome == Infrastructure.StudioOutcome.Cancelled;
    public bool NeedsChecking => Outcome == Infrastructure.StudioOutcome.NeedsChecking;

    private void NotifyOutcome() {
        foreach (string name in new[] { nameof(IsSucceeded), nameof(HasNotes), nameof(IsFailed), nameof(IsCancelled), nameof(NeedsChecking) }) OnPropertyChanged(name);
    }

    /// <summary>Completed, but the conversion's fidelity evidence reports approximated or omitted content.</summary>
    public bool HasLoss => State == ConversionJobState.Completed && ConversionEvidence?.HasLoss == true;

    internal bool CanRetry => State is ConversionJobState.Failed or ConversionJobState.Cancelled;

    internal void PrepareAttempt() {
        State = ConversionJobState.Queued;
        Status = T("Queued", "Queued");
        ProgressFraction = 0D;
        OutputPath = null;
        Summary = null;
        Diagnostics = Array.Empty<OfficeWorkflowDiagnostic>();
        ConversionEvidence = null;
        OnPropertyChanged(nameof(SourceFingerprint));
        OnPropertyChanged(nameof(HasWarnings));
        OnPropertyChanged(nameof(HasLoss));
    }

    internal void ReportProgress(OfficeWorkflowProgress progress) {
        if (State is not (ConversionJobState.Queued or ConversionJobState.Running)) return;
        State = ConversionJobState.Running;
        Status = _localizer.FormatOrDefault("Conversion.Job.Running", "Running · {0}", progress.Stage.Replace('-', ' '));
        ProgressFraction = progress.Fraction;
    }

    internal void EndWithoutResult(bool cancelled, string message) {
        State = cancelled ? ConversionJobState.Cancelled : ConversionJobState.Unconfirmed;
        Status = cancelled ? T("Cancelled", "Cancelled") : T("Unconfirmed", "Check output");
        Summary = message;
    }

    internal void FailBeforeExecution(string message) {
        State = ConversionJobState.Failed;
        Status = T("Failed", "Failed");
        Summary = message;
    }

    internal void Apply(OfficeWorkflowResult result) {
        State = result.Status switch {
            OfficeWorkflowStatus.Completed => ConversionJobState.Completed,
            OfficeWorkflowStatus.Cancelled => ConversionJobState.Cancelled,
            OfficeWorkflowStatus.Unconfirmed => ConversionJobState.Unconfirmed,
            _ => ConversionJobState.Failed
        };
        OutputPath = result.OutputPath;
        Summary = result.Summary;
        Diagnostics = result.Diagnostics;
        ConversionEvidence = result.ConversionEvidence;
        OnPropertyChanged(nameof(SourceFingerprint));
        ProgressFraction = result.Status == OfficeWorkflowStatus.Cancelled ? ProgressFraction : 1D;
        Status = result.Status switch {
            OfficeWorkflowStatus.Completed when HasWarnings || HasLoss => T("CompletedWithWarnings", "Completed with warnings"),
            OfficeWorkflowStatus.Completed => T("Completed", "Completed"),
            OfficeWorkflowStatus.Cancelled => T("Cancelled", "Cancelled"),
            OfficeWorkflowStatus.Unconfirmed => T("Unconfirmed", "Check output"),
            _ => T("Failed", "Failed")
        };
        OnPropertyChanged(nameof(HasWarnings));
        OnPropertyChanged(nameof(HasLoss));
        NotifyOutcome();
    }

    private string T(string suffix, string fallback) =>
        _localizer.GetOrDefault("Conversion.Job." + suffix, fallback);
}
