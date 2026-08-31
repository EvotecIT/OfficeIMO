using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed record WorkflowProfileChoice(OfficeWorkflowOutputProfile Value, string Label, string Description);

public sealed record WorkflowConflictChoice(OfficeWorkflowConflictPolicy Value, string Label, string Description);

public sealed record HealthOperationChoice(OfficeWorkflowOperation Value, string Label, string Description, bool ProducesArtifact);

public sealed class ConversionRouteChoice {
    public ConversionRouteChoice(OfficeWorkflowRoute route) => Route = route;

    public OfficeWorkflowRoute Route { get; }
    public string Label => Route.Label;
    public string Description => Route.Description;
    public string Engine => Route.Engine;
    public string Fidelity => Route.Fidelity + " · " + Route.SupportLevel;
    public string KnownLimitations => Route.KnownLimitations;
}
