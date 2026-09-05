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

    internal ConversionJobViewModel(string inputPath, ConversionRouteChoice route, IStudioLocalizer? localizer) {
        Id = Guid.NewGuid().ToString("N");
        InputPath = inputPath;
        Route = route;
        _localizer = localizer ?? StudioLocalization.Current;
        Status = _localizer.GetOrDefault("Conversion.Job.Queued", "Queued");
    }

    public string Id { get; }
    public string InputPath { get; }
    public string FileName => Path.GetFileName(InputPath);
    public ConversionRouteChoice Route { get; }
    public string RouteLabel => Route.Route.Source + " → " + Route.Route.Target;
    public string Engine => Route.Engine;
    public string Fidelity => Route.Fidelity;
    public string KnownLimitations => Route.KnownLimitations;

    [ObservableProperty]
    private string _status = string.Empty;

    [ObservableProperty]
    private double _progressFraction;

    [ObservableProperty]
    private string? _outputPath;

    [ObservableProperty]
    private string? _summary;

    [ObservableProperty]
    private IReadOnlyList<OfficeWorkflowDiagnostic> _diagnostics = Array.Empty<OfficeWorkflowDiagnostic>();

    public bool HasWarnings => Diagnostics.Any(item => item.Severity == OfficeWorkflowDiagnosticSeverity.Warning);

    [ObservableProperty]
    private ConversionJobState _state;

    internal bool CanRetry => State is ConversionJobState.Failed or ConversionJobState.Cancelled;

    internal void PrepareAttempt() {
        State = ConversionJobState.Queued;
        Status = T("Queued", "Queued");
        ProgressFraction = 0D;
        OutputPath = null;
        Summary = null;
        Diagnostics = Array.Empty<OfficeWorkflowDiagnostic>();
        OnPropertyChanged(nameof(HasWarnings));
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

    internal void Apply(OfficeWorkflowResult result) {
        State = result.Status switch {
            OfficeWorkflowStatus.Completed => ConversionJobState.Completed,
            OfficeWorkflowStatus.Cancelled => ConversionJobState.Cancelled,
            _ => ConversionJobState.Failed
        };
        OutputPath = result.OutputPath;
        Summary = result.Summary;
        Diagnostics = result.Diagnostics;
        ProgressFraction = result.Status == OfficeWorkflowStatus.Cancelled ? ProgressFraction : 1D;
        Status = result.Status switch {
            OfficeWorkflowStatus.Completed when HasWarnings => T("CompletedWithWarnings", "Completed with warnings"),
            OfficeWorkflowStatus.Completed => T("Completed", "Completed"),
            OfficeWorkflowStatus.Cancelled => T("Cancelled", "Cancelled"),
            _ => T("Failed", "Failed")
        };
        OnPropertyChanged(nameof(HasWarnings));
    }

    private string T(string suffix, string fallback) =>
        _localizer.GetOrDefault("Conversion.Job." + suffix, fallback);
}
