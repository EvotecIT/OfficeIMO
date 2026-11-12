namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private readonly IReadOnlyDictionary<string, OfficeWorkflowConversionRegistration> _conversions;
    private readonly Func<ValidatedRequest, List<OfficeWorkflowDiagnostic>, CancellationToken, OperationArtifact> _execute;
    private readonly bool _registeredConversionsOnly;

    /// <summary>Creates a runner whose requests run only the supplied conversions.</summary>
    /// <remarks>
    /// Requests for built-in routes and non-conversion operations fail validation. Source capture, limits,
    /// output reopening and atomic publication match the full runner. NativeAOT applications that use this
    /// runner compile only these converters, not every built-in format.
    /// </remarks>
    public static OfficeWorkflowRunner ForRegisteredConversions(IEnumerable<OfficeWorkflowConversionRegistration> conversions) {
        ArgumentNullException.ThrowIfNull(conversions);
        return new OfficeWorkflowRunner(null, null, conversions, ExecuteRegisteredConversion, includeBuiltInOperations: false);
    }

    private static OperationArtifact ExecuteRegisteredConversion(
        ValidatedRequest request,
        List<OfficeWorkflowDiagnostic> diagnostics,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        byte[] input = OfficeWorkflowInputReader.ReadAllBytes(request.InputPath, request.Limits.MaximumInputBytes, cancellationToken);
        return ConvertRegistered(request, input, diagnostics, cancellationToken);
    }

    private static (IReadOnlyDictionary<string, OfficeWorkflowConversionRegistration>, IReadOnlyList<OfficeWorkflowRoute>)
        SnapshotRegistrations(IEnumerable<OfficeWorkflowConversionRegistration> conversions, bool includeBuiltInRoutes) {
        ArgumentNullException.ThrowIfNull(conversions);
        var registrations = new Dictionary<string, OfficeWorkflowConversionRegistration>(StringComparer.OrdinalIgnoreCase);
        foreach (OfficeWorkflowConversionRegistration conversion in conversions) {
            ArgumentNullException.ThrowIfNull(conversion);
            if (!registrations.TryAdd(conversion.RouteId, conversion))
                throw new ArgumentException("A conversion route was registered more than once: " + conversion.RouteId, nameof(conversions));
        }
        IReadOnlyList<OfficeWorkflowRoute> routes = Array.AsReadOnly(OfficeConversionCapabilityCatalog.All
            .Where(capability => registrations.ContainsKey(capability.Id)
                || (includeBuiltInRoutes && OfficeWorkflowCatalog.FindExecutable(capability.Id) is not null))
            .Select(capability => new OfficeWorkflowRoute(capability, canExecute: true))
            .OrderBy(route => route.Source, StringComparer.Ordinal).ThenBy(route => route.Target, StringComparer.Ordinal).ToArray());
        return (registrations, routes);
    }

    /// <summary>Executable routes for this configured runner, including opt-in registrations.</summary>
    public IReadOnlyList<OfficeWorkflowRoute> ConversionRoutes { get; }

    /// <summary>Reports an actual configured route's directory-package owner without reading or capturing that package.</summary>
    internal bool SupportsDirectoryPackageInput(string routeId) =>
        _conversions.TryGetValue(routeId, out OfficeWorkflowConversionRegistration? registration) && registration.DirectoryPackageInput is not null;

    private static OperationArtifact ConvertRegistered(ValidatedRequest request, byte[] input,
        List<OfficeWorkflowDiagnostic> diagnostics, CancellationToken token) {
        using var source = new MemoryStream(input, writable: false);
        using var output = new OfficeWorkflowBoundedMemoryStream(request.Limits.MaximumOutputBytes);
        OfficeWorkflowConversionEvidence evidence;
        try {
            evidence = request.Registration!.Converter(source, output, request.Limits.CloneAndValidate(), request.RegisteredConversionSettings, token)
                ?? throw new InvalidOperationException("The registered converter returned no evidence.");
        } catch (OfficeConversionException exception) {
            var failed = exception.Report as OfficeWorkflowConversionEvidence ?? new OfficeWorkflowConversionEvidence(exception.Report);
            if (token.IsCancellationRequested && exception.InnerException is OperationCanceledException cancelled)
                throw new WorkflowConversionCancellationException(cancelled, failed);
            throw new WorkflowConversionFailureException(exception.InnerException ?? exception, failed);
        }
        AddConversionDiagnostics(evidence, diagnostics);
        diagnostics.Add(new OfficeWorkflowDiagnostic("ConversionEvidence", "Source and projection evidence from the registered format owner.",
            stage: "convert", details: evidence.Facts));
        diagnostics.Add(new OfficeWorkflowDiagnostic("RouteContract", request.Route!.Description, stage: "convert",
            details: new Dictionary<string, string> { ["route"] = request.Route.Id, ["engine"] = request.Route.Engine,
                ["knownLimitations"] = request.Route.KnownLimitations }));
        try {
            byte[] bytes = output.ToArray();
            if (request.ConversionOptions?.CompressPdfOutput == true)
                bytes = CompressConvertedPdfOutput(bytes, request, diagnostics, token);
            token.ThrowIfCancellationRequested();
            return new OperationArtifact(bytes, request.Route.Label + (evidence.HasLoss
                ? " completed with fidelity warnings; review the structured diagnostics."
                : " completed and the output reopened successfully."), null, ConversionEvidence: evidence);
        } catch (OperationCanceledException exception) when (token.IsCancellationRequested) {
            throw new WorkflowConversionCancellationException(exception, evidence, diagnosticsAdded: true);
        } catch (Exception exception) when (exception is not OperationCanceledException and not OutOfMemoryException and not StackOverflowException) {
            throw new WorkflowConversionFailureException(exception, evidence, diagnosticsAdded: true);
        }
    }
}
