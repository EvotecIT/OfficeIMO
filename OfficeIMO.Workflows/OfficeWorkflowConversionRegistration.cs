using System.Collections.ObjectModel;

namespace OfficeIMO.Workflows;

// Carries an owning converter's evidence across the runner's structured failure boundary.
internal sealed class WorkflowConversionFailureException : InvalidOperationException {
    internal WorkflowConversionFailureException(Exception cause, OfficeWorkflowConversionEvidence evidence, bool diagnosticsAdded = false,
        ISet<string>? importDiagnosticSources = null)
        : base(cause.Message, cause) {
        Evidence = evidence;
        DiagnosticsAdded = diagnosticsAdded;
        ImportDiagnosticSources = importDiagnosticSources == null ? null : new HashSet<string>(importDiagnosticSources, StringComparer.Ordinal);
    }

    internal OfficeWorkflowConversionEvidence Evidence { get; }
    internal bool DiagnosticsAdded { get; }
    internal ISet<string>? ImportDiagnosticSources { get; }
}

// Retains available reports while preserving the runner's cancellation classification.
internal sealed class WorkflowConversionCancellationException : OperationCanceledException {
    internal WorkflowConversionCancellationException(OperationCanceledException cause, OfficeWorkflowConversionEvidence evidence,
        bool diagnosticsAdded = false, ISet<string>? importDiagnosticSources = null) : base(cause.Message, cause, cause.CancellationToken) {
        Evidence = evidence;
        DiagnosticsAdded = diagnosticsAdded;
        ImportDiagnosticSources = importDiagnosticSources == null ? null : new HashSet<string>(importDiagnosticSources, StringComparer.Ordinal);
    }
    internal OfficeWorkflowConversionEvidence Evidence { get; }
    internal bool DiagnosticsAdded { get; }
    internal ISet<string>? ImportDiagnosticSources { get; }
}

/// <summary>Converts a captured input into a bounded output stream. Streams remain owned by the runner.</summary>
/// <remarks>The delegate must honor cancellation and leave both streams open. Publication and output reopen validation belong to the runner.
/// Throw <see cref="OfficeConversionException"/> with the available report to retain evidence when output serialization fails.
/// An <see cref="OfficeWorkflowConversionEvidence"/> report also retains its compact facts; an inner cancellation exception preserves cancellation classification.</remarks>
public delegate OfficeWorkflowConversionEvidence OfficeWorkflowConverter(
    Stream input, Stream output, OfficeOperationLimits limits, CancellationToken cancellationToken);

/// <summary>An opt-in implementation of an existing canonical conversion route.</summary>
public sealed class OfficeWorkflowConversionRegistration {
    /// <summary>Registers a converter without replacing a built-in route or creating another capability catalog.</summary>
    public OfficeWorkflowConversionRegistration(string routeId, OfficeWorkflowConverter converter)
        : this(routeId, Adapt(converter), null) { }

    private OfficeWorkflowConversionRegistration(string routeId,
        Func<Stream, Stream, OfficeOperationLimits, IOfficeConversionSettings?, CancellationToken, OfficeWorkflowConversionEvidence> converter,
        Func<IOfficeConversionSettings, bool>? acceptsSettings,
        Func<string, OfficeOperationLimits, IOfficeConversionSettings?, OfficeWorkflowStreamInput>? directoryPackageInput = null) {
        OfficeWorkflowRoute route = OfficeWorkflowCatalog.Find(routeId)
            ?? throw new ArgumentException("Choose an existing canonical conversion route.", nameof(routeId));
        if (route.CanExecute) throw new ArgumentException("Built-in conversion routes cannot be replaced.", nameof(routeId));
        if (route.TargetExtension.TrimStart('.').ToLowerInvariant() is not ("docx" or "xlsx" or "pptx" or "pdf"))
            throw new NotSupportedException("Opt-in workflow conversion supports DOCX, XLSX, PPTX, and PDF destinations.");
        RouteId = route.Id;
        Converter = converter;
        _acceptsSettings = acceptsSettings;
        DirectoryPackageInput = directoryPackageInput;
    }

    /// <summary>Canonical capability identifier.</summary>
    public string RouteId { get; }
    internal Func<Stream, Stream, OfficeOperationLimits, IOfficeConversionSettings?, CancellationToken, OfficeWorkflowConversionEvidence> Converter { get; }
    private readonly Func<IOfficeConversionSettings, bool>? _acceptsSettings;
    internal Func<string, OfficeOperationLimits, IOfficeConversionSettings?, OfficeWorkflowStreamInput>? DirectoryPackageInput { get; }

    /// <summary>Registers a canonical route with an explicit adapter-owned settings contract.</summary>
    public static OfficeWorkflowConversionRegistration Create<TSettings>(string routeId,
        OfficeWorkflowConfiguredConverter<TSettings> converter) where TSettings : class, IOfficeConversionSettings =>
        Create(routeId, converter, null);

    /// <summary>Registers a canonical route with an explicit adapter-owned settings contract.</summary>
    /// <param name="routeId">Canonical conversion route.</param>
    /// <param name="converter">Conversion owner.</param>
    /// <param name="directoryPackageInput">Optional local directory-package owner. It must create a bounded, repeatable snapshot stream, verify membership/content on every reopen, and preserve source identity. Called after request settings are validated.</param>
    public static OfficeWorkflowConversionRegistration Create<TSettings>(string routeId,
        OfficeWorkflowConfiguredConverter<TSettings> converter,
        Func<string, OfficeOperationLimits, TSettings?, OfficeWorkflowStreamInput>? directoryPackageInput) where TSettings : class, IOfficeConversionSettings {
        if (converter == null) throw new ArgumentNullException(nameof(converter));
        return new(routeId, (input, output, limits, settings, token) => converter(input, output, limits, (TSettings?)settings, token),
            settings => settings is TSettings, directoryPackageInput is null ? null
                : (path, limits, settings) => directoryPackageInput(path, limits, (TSettings?)settings));
    }

    internal IOfficeConversionSettings? SnapshotSettings(IOfficeConversionSettings? settings) {
        if (settings is null) return null;
        if (_acceptsSettings?.Invoke(settings) != true)
            throw new ArgumentException("The registered conversion route does not accept these settings.");
        IOfficeConversionSettings snapshot = settings.Snapshot()
            ?? throw new ArgumentException("The conversion settings returned no snapshot.");
        if (!_acceptsSettings(snapshot)) throw new ArgumentException("The conversion settings snapshot changed its contract type.");
        return snapshot;
    }

    private static Func<Stream, Stream, OfficeOperationLimits, IOfficeConversionSettings?, CancellationToken, OfficeWorkflowConversionEvidence>
        Adapt(OfficeWorkflowConverter converter) {
        if (converter == null) throw new ArgumentNullException(nameof(converter));
        return (input, output, limits, _, token) => converter(input, output, limits, token);
    }
}

/// <summary>Immutable conversion evidence retained independently of the destination document lifetime.</summary>
public sealed class OfficeWorkflowConversionEvidence : IOfficeConversionReport {
    /// <summary>Copies typed fidelity diagnostics and compact source/projection facts from the owning converter.</summary>
    public OfficeWorkflowConversionEvidence(IOfficeConversionReport report, IReadOnlyDictionary<string, string>? facts = null) {
        if (report == null) throw new ArgumentNullException(nameof(report));
        FidelityDiagnostics = OfficeConversionFidelityDiagnostics.Flatten([report]);
        Facts = new ReadOnlyDictionary<string, string>((facts ?? new Dictionary<string, string>()).ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal));
    }

    /// <summary>Copies evidence from ordered conversion stages after their output operations have completed.</summary>
    public OfficeWorkflowConversionEvidence(IReadOnlyList<IOfficeConversionReport> reports, IReadOnlyDictionary<string, string> facts) {
        if (reports == null) throw new ArgumentNullException(nameof(reports));
        if (facts == null) throw new ArgumentNullException(nameof(facts));
        FidelityDiagnostics = OfficeConversionFidelityDiagnostics.Flatten(reports);
        Facts = new ReadOnlyDictionary<string, string>(facts.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal));
    }

    /// <inheritdoc />
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; }
    /// <summary>Compact producer, projection, coverage, and reconstruction evidence. These facts do not establish visual equivalence.</summary>
    public IReadOnlyDictionary<string, string> Facts { get; }
    /// <inheritdoc />
    public bool HasLoss => FidelityDiagnostics.Any(diagnostic => diagnostic.LossKind != OfficeConversionLossKind.None);
    /// <inheritdoc />
    public void RequireNoLoss() {
        if (HasLoss) throw new InvalidOperationException("The workflow conversion reported fidelity loss or unassessed content.");
    }
}
