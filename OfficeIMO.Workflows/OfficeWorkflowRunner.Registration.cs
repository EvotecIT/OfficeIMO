namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private readonly IReadOnlyDictionary<string, OfficeWorkflowConversionRegistration> _conversions;

    private static (IReadOnlyDictionary<string, OfficeWorkflowConversionRegistration>, IReadOnlyList<OfficeWorkflowRoute>)
        SnapshotRegistrations(IEnumerable<OfficeWorkflowConversionRegistration> conversions) {
        ArgumentNullException.ThrowIfNull(conversions);
        var registrations = new Dictionary<string, OfficeWorkflowConversionRegistration>(StringComparer.OrdinalIgnoreCase);
        foreach (OfficeWorkflowConversionRegistration conversion in conversions) {
            ArgumentNullException.ThrowIfNull(conversion);
            if (!registrations.TryAdd(conversion.RouteId, conversion))
                throw new ArgumentException("A conversion route was registered more than once: " + conversion.RouteId, nameof(conversions));
        }
        IReadOnlyList<OfficeWorkflowRoute> routes = Array.AsReadOnly(OfficeConversionCapabilityCatalog.All
            .Where(capability => OfficeWorkflowCatalog.FindExecutable(capability.Id) is not null || registrations.ContainsKey(capability.Id))
            .Select(capability => new OfficeWorkflowRoute(capability, canExecute: true))
            .OrderBy(route => route.Source, StringComparer.Ordinal).ThenBy(route => route.Target, StringComparer.Ordinal).ToArray());
        return (registrations, routes);
    }

    /// <summary>Executable routes for this configured runner, including opt-in registrations.</summary>
    public IReadOnlyList<OfficeWorkflowRoute> ConversionRoutes { get; }

    private static OperationArtifact ConvertRegistered(ValidatedRequest request, byte[] input,
        List<OfficeWorkflowDiagnostic> diagnostics, CancellationToken token) {
        using var source = new MemoryStream(input, writable: false);
        using var output = new OfficeWorkflowBoundedMemoryStream(request.Limits.MaximumOutputBytes);
        OfficeWorkflowConversionEvidence evidence = request.Registration!.Converter(source, output, request.Limits.CloneAndValidate(), token)
            ?? throw new InvalidOperationException("The registered converter returned no evidence.");
        token.ThrowIfCancellationRequested();
        foreach (OfficeConversionFidelityDiagnostic diagnostic in evidence.FidelityDiagnostics) {
            var details = new Dictionary<string, string>(StringComparer.Ordinal) {
                ["lossKind"] = diagnostic.LossKind.ToString(), ["source"] = diagnostic.Source
            };
            if (diagnostic.Location is not null) details["location"] = diagnostic.Location;
            diagnostics.Add(new OfficeWorkflowDiagnostic(diagnostic.Code, diagnostic.Message,
                diagnostic.LossKind == OfficeConversionLossKind.Failure ? OfficeWorkflowDiagnosticSeverity.Error
                    : diagnostic.LossKind == OfficeConversionLossKind.None ? OfficeWorkflowDiagnosticSeverity.Information
                    : OfficeWorkflowDiagnosticSeverity.Warning, "convert", details));
        }
        diagnostics.Add(new OfficeWorkflowDiagnostic("ConversionEvidence", "Source and projection evidence from the registered format owner.",
            stage: "convert", details: evidence.Facts));
        diagnostics.Add(new OfficeWorkflowDiagnostic("RouteContract", request.Route!.Description, stage: "convert",
            details: new Dictionary<string, string> { ["route"] = request.Route.Id, ["engine"] = request.Route.Engine,
                ["knownLimitations"] = request.Route.KnownLimitations }));
        return new OperationArtifact(output.ToArray(), request.Route.Label + (evidence.HasLoss
            ? " completed with fidelity warnings; review the structured diagnostics."
            : " completed and the output reopened successfully."), null, ConversionEvidence: evidence);
    }
}
