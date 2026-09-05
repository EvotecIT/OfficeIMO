using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Workflows;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Workflows;

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

    internal void Apply(OfficeWorkflowResult result) {
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
