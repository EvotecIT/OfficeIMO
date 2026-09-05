namespace OfficeIMO.Workflows;

/// <summary>Creates fluent local document workflows over <see cref="OfficeWorkflowRunner"/>.</summary>
public static class OfficeWorkflow {
    /// <summary>Creates a conversion workflow. Supply a route with <see cref="OfficeWorkflowBuilder.Via"/> or an output path with <see cref="OfficeWorkflowBuilder.To"/>.</summary>
    public static OfficeWorkflowBuilder Convert(string inputPath) =>
        Create(OfficeWorkflowOperation.Convert, inputPath);

    /// <summary>Creates a PDF inspection workflow.</summary>
    public static OfficeWorkflowBuilder Inspect(string inputPath) =>
        Create(OfficeWorkflowOperation.Inspect, inputPath);

    /// <summary>Creates a PDF comparison workflow.</summary>
    public static OfficeWorkflowBuilder Compare(string inputPath, string comparisonPath) =>
        Create(OfficeWorkflowOperation.Compare, inputPath).Against(comparisonPath);

    /// <summary>Creates a lossless PDF optimization workflow.</summary>
    public static OfficeWorkflowBuilder Optimize(string inputPath) =>
        Create(OfficeWorkflowOperation.Optimize, inputPath);

    /// <summary>Creates a report-only PDF repair-planning workflow.</summary>
    public static OfficeWorkflowBuilder PlanRepair(string inputPath) =>
        Create(OfficeWorkflowOperation.RepairPlan, inputPath);

    /// <summary>Creates a verified PDF repair workflow.</summary>
    public static OfficeWorkflowBuilder Repair(string inputPath) =>
        Create(OfficeWorkflowOperation.Repair, inputPath);

    /// <summary>Creates a verified PDF sanitization workflow.</summary>
    public static OfficeWorkflowBuilder Sanitize(string inputPath) =>
        Create(OfficeWorkflowOperation.Sanitize, inputPath);

    /// <summary>Runs an explicitly constructed request through the default local runner.</summary>
    public static Task<OfficeWorkflowResult> RunAsync(
        OfficeWorkflowRequest request,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default) =>
        new OfficeWorkflowRunner().RunAsync(request, progress, cancellationToken);

    private static OfficeWorkflowBuilder Create(OfficeWorkflowOperation operation, string inputPath) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("Input path cannot be empty.", nameof(inputPath));
        return new OfficeWorkflowBuilder(new OfficeWorkflowRequest {
            Operation = operation,
            InputPath = inputPath
        });
    }
}

/// <summary>Fluent configuration for one local document workflow.</summary>
public sealed class OfficeWorkflowBuilder {
    private readonly OfficeWorkflowRequest _request;

    internal OfficeWorkflowBuilder(OfficeWorkflowRequest request) {
        _request = request ?? throw new ArgumentNullException(nameof(request));
    }

    /// <summary>Sets the conversion route identifier from <see cref="OfficeWorkflowCatalog"/>.</summary>
    public OfficeWorkflowBuilder Via(string routeId) {
        if (string.IsNullOrWhiteSpace(routeId)) throw new ArgumentException("Conversion route id cannot be empty.", nameof(routeId));
        _request.ConversionRouteId = routeId;
        return this;
    }

    /// <summary>Sets the output artifact path.</summary>
    public OfficeWorkflowBuilder To(string outputPath) {
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("Output path cannot be empty.", nameof(outputPath));
        _request.OutputPath = outputPath;
        return this;
    }

    /// <summary>Sets the comparison input for a compare workflow.</summary>
    public OfficeWorkflowBuilder Against(string comparisonPath) {
        if (string.IsNullOrWhiteSpace(comparisonPath)) throw new ArgumentException("Comparison path cannot be empty.", nameof(comparisonPath));
        _request.ComparisonPath = comparisonPath;
        return this;
    }

    /// <summary>Sets the cross-format output profile.</summary>
    public OfficeWorkflowBuilder WithProfile(OfficeWorkflowOutputProfile profile) {
        _request.OutputProfile = profile;
        return this;
    }

    /// <summary>Sets destination conflict behavior.</summary>
    public OfficeWorkflowBuilder OnConflict(OfficeWorkflowConflictPolicy policy) {
        _request.ConflictPolicy = policy;
        return this;
    }

    /// <summary>Sets the PDF password used for the primary input.</summary>
    public OfficeWorkflowBuilder WithPdfPassword(string? password) {
        _request.PdfPassword = password;
        return this;
    }

    /// <summary>Sets the PDF password used for the comparison input.</summary>
    public OfficeWorkflowBuilder WithComparisonPdfPassword(string? password) {
        _request.ComparisonPdfPassword = password;
        return this;
    }

    /// <summary>Sets bounded input and output sizes for the workflow.</summary>
    public OfficeWorkflowBuilder WithLimits(long maximumInputBytes, long maximumOutputBytes) {
        _request.Limits = new OfficeWorkflowLimits {
            MaximumInputBytes = maximumInputBytes,
            MaximumOutputBytes = maximumOutputBytes
        };
        return this;
    }

    /// <summary>Sets the caller-visible request identifier.</summary>
    public OfficeWorkflowBuilder WithId(string id) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Request id cannot be empty.", nameof(id));
        _request.Id = id;
        return this;
    }

    /// <summary>Builds an independent request snapshot, inferring a locally executable conversion route from file extensions when possible.</summary>
    public OfficeWorkflowRequest Build() {
        string? routeId = _request.ConversionRouteId;
        if (_request.Operation == OfficeWorkflowOperation.Convert && string.IsNullOrWhiteSpace(routeId)) {
            if (string.IsNullOrWhiteSpace(_request.OutputPath)) {
                throw new InvalidOperationException("A conversion workflow requires Via(routeId) or To(outputPath) so the route can be selected.");
            }

            OfficeWorkflowRoute? route = OfficeWorkflowCatalog.Find(
                Path.GetExtension(_request.InputPath),
                Path.GetExtension(_request.OutputPath),
                executableOnly: false);
            routeId = route is { CanExecute: true }
                ? route.Id
                : throw new NotSupportedException(
                    "No unique locally executable conversion route matches the input and output extensions. Use Via(routeId) when formats share an extension.");
        }

        OfficeWorkflowLimits limits = _request.Limits ?? throw new InvalidOperationException("Workflow limits cannot be null.");
        return new OfficeWorkflowRequest {
            Id = _request.Id,
            Operation = _request.Operation,
            InputPath = _request.InputPath,
            ComparisonPath = _request.ComparisonPath,
            ConversionRouteId = routeId,
            OutputPath = _request.OutputPath,
            ConflictPolicy = _request.ConflictPolicy,
            OutputProfile = _request.OutputProfile,
            PdfPassword = _request.PdfPassword,
            ComparisonPdfPassword = _request.ComparisonPdfPassword,
            Limits = new OfficeWorkflowLimits {
                MaximumInputBytes = limits.MaximumInputBytes,
                MaximumOutputBytes = limits.MaximumOutputBytes
            }
        };
    }

    /// <summary>Runs this workflow through the supplied runner or the default local runner.</summary>
    public Task<OfficeWorkflowResult> RunAsync(
        IOfficeWorkflowRunner? runner = null,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default) =>
        (runner ?? new OfficeWorkflowRunner()).RunAsync(Build(), progress, cancellationToken);
}
