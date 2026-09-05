using OfficeIMO.Workflows;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed record WorkflowProfileChoice(OfficeWorkflowOutputProfile Value, string Label, string Description);

public sealed record WorkflowConflictChoice(OfficeWorkflowConflictPolicy Value, string Label, string Description);

public sealed record HealthOperationChoice(OfficeWorkflowOperation Value, string Label, string Description, bool ProducesArtifact);

public sealed class ConversionRouteChoice {
    private readonly IStudioLocalizer? _localizer;

    public ConversionRouteChoice(OfficeWorkflowRoute route) : this(route, null) { }

    internal ConversionRouteChoice(OfficeWorkflowRoute route, IStudioLocalizer? localizer) {
        Route = route;
        _localizer = localizer;
    }

    public OfficeWorkflowRoute Route { get; }
    public string Label => Localize("Label", Route.Label);
    public string Description => Localize("Description", Route.Description);
    public string Engine => Route.Engine;
    public string Fidelity => Localize("Fidelity", Route.Fidelity) + " · " + Localize("SupportLevel", Route.SupportLevel);
    public string KnownLimitations => Localize("KnownLimitations", Route.KnownLimitations);

    private string Localize(string property, string fallback) =>
        _localizer?.GetOrDefault($"Conversion.Route.{Route.Id}.{property}", fallback) ?? fallback;
}
