using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class ConversionJobViewModel : ObservableObject {
    public ConversionJobViewModel(string inputPath, ConversionRouteChoice route) {
        Id = Guid.NewGuid().ToString("N");
        InputPath = inputPath;
        Route = route;
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
    private string _status = "Queued";

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
            OfficeWorkflowStatus.Completed when HasWarnings => "Completed with warnings",
            OfficeWorkflowStatus.Completed => "Completed",
            OfficeWorkflowStatus.Cancelled => "Cancelled",
            _ => "Failed"
        };
        OnPropertyChanged(nameof(HasWarnings));
    }
}
