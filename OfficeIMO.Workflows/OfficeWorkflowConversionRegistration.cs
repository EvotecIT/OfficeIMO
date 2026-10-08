using System.Collections.ObjectModel;

namespace OfficeIMO.Workflows;

// Carries an owning converter's evidence across the runner's structured failure boundary.
internal sealed class WorkflowConversionFailureException : InvalidOperationException {
    internal WorkflowConversionFailureException(Exception cause, OfficeWorkflowConversionEvidence evidence)
        : base(cause.Message, cause) => Evidence = evidence;

    internal OfficeWorkflowConversionEvidence Evidence { get; }
}

/// <summary>Converts a captured input into a bounded output stream. Streams remain owned by the runner.</summary>
/// <remarks>The delegate must honor cancellation and leave both streams open. Publication and output reopen validation belong to the runner.</remarks>
public delegate OfficeWorkflowConversionEvidence OfficeWorkflowConverter(
    Stream input, Stream output, OfficeWorkflowLimits limits, CancellationToken cancellationToken);

/// <summary>An opt-in implementation of an existing canonical conversion route.</summary>
public sealed class OfficeWorkflowConversionRegistration {
    /// <summary>Registers a converter without replacing a built-in route or creating another capability catalog.</summary>
    public OfficeWorkflowConversionRegistration(string routeId, OfficeWorkflowConverter converter)
        : this(routeId, Adapt(converter), null) { }

    private OfficeWorkflowConversionRegistration(string routeId,
        Func<Stream, Stream, OfficeWorkflowLimits, IOfficeWorkflowConversionSettings?, CancellationToken, OfficeWorkflowConversionEvidence> converter,
        Func<IOfficeWorkflowConversionSettings, bool>? acceptsSettings,
        Func<string, OfficeWorkflowLimits, IOfficeWorkflowConversionSettings?, OfficeWorkflowStreamInput>? directoryPackageInput = null) {
        OfficeWorkflowRoute route = OfficeWorkflowCatalog.Find(routeId)
            ?? throw new ArgumentException("Choose an existing canonical conversion route.", nameof(routeId));
        if (route.CanExecute) throw new ArgumentException("Built-in conversion routes cannot be replaced.", nameof(routeId));
        if (route.TargetExtension.TrimStart('.').ToLowerInvariant() is not ("docx" or "xlsx" or "pptx"))
            throw new NotSupportedException("Opt-in workflow conversion currently supports DOCX, XLSX, and PPTX destinations.");
        RouteId = route.Id;
        Converter = converter;
        _acceptsSettings = acceptsSettings;
        DirectoryPackageInput = directoryPackageInput;
    }

    /// <summary>Canonical capability identifier.</summary>
    public string RouteId { get; }
    internal Func<Stream, Stream, OfficeWorkflowLimits, IOfficeWorkflowConversionSettings?, CancellationToken, OfficeWorkflowConversionEvidence> Converter { get; }
    private readonly Func<IOfficeWorkflowConversionSettings, bool>? _acceptsSettings;
    internal Func<string, OfficeWorkflowLimits, IOfficeWorkflowConversionSettings?, OfficeWorkflowStreamInput>? DirectoryPackageInput { get; }

    /// <summary>Registers a canonical route with an explicit adapter-owned settings contract.</summary>
    public static OfficeWorkflowConversionRegistration Create<TSettings>(string routeId,
        OfficeWorkflowConfiguredConverter<TSettings> converter) where TSettings : class, IOfficeWorkflowConversionSettings =>
        Create(routeId, converter, null);

    /// <summary>Registers a canonical route with an explicit adapter-owned settings contract.</summary>
    /// <param name="routeId">Canonical conversion route.</param>
    /// <param name="converter">Conversion owner.</param>
    /// <param name="directoryPackageInput">Optional local directory-package owner. It must create a bounded, repeatable snapshot stream, verify membership/content on every reopen, and preserve source identity. Called after request settings are validated.</param>
    public static OfficeWorkflowConversionRegistration Create<TSettings>(string routeId,
        OfficeWorkflowConfiguredConverter<TSettings> converter,
        Func<string, OfficeWorkflowLimits, TSettings?, OfficeWorkflowStreamInput>? directoryPackageInput) where TSettings : class, IOfficeWorkflowConversionSettings {
        ArgumentNullException.ThrowIfNull(converter);
        return new(routeId, (input, output, limits, settings, token) => converter(input, output, limits, (TSettings?)settings, token),
            settings => settings is TSettings, directoryPackageInput is null ? null
                : (path, limits, settings) => directoryPackageInput(path, limits, (TSettings?)settings));
    }

    internal IOfficeWorkflowConversionSettings? SnapshotSettings(IOfficeWorkflowConversionSettings? settings) {
        if (settings is null) return null;
        if (_acceptsSettings?.Invoke(settings) != true)
            throw new ArgumentException("The registered conversion route does not accept these settings.");
        IOfficeWorkflowConversionSettings snapshot = settings.Snapshot()
            ?? throw new ArgumentException("The conversion settings returned no snapshot.");
        if (!_acceptsSettings(snapshot)) throw new ArgumentException("The conversion settings snapshot changed its contract type.");
        return snapshot;
    }

    private static Func<Stream, Stream, OfficeWorkflowLimits, IOfficeWorkflowConversionSettings?, CancellationToken, OfficeWorkflowConversionEvidence>
        Adapt(OfficeWorkflowConverter converter) {
        ArgumentNullException.ThrowIfNull(converter);
        return (input, output, limits, _, token) => converter(input, output, limits, token);
    }
}

/// <summary>Immutable conversion evidence retained independently of the destination document lifetime.</summary>
public sealed class OfficeWorkflowConversionEvidence : IOfficeConversionReport {
    /// <summary>Copies typed fidelity diagnostics and compact source/projection facts from the owning converter.</summary>
    public OfficeWorkflowConversionEvidence(IOfficeConversionReport report, IReadOnlyDictionary<string, string>? facts = null) {
        ArgumentNullException.ThrowIfNull(report);
        FidelityDiagnostics = OfficeConversionFidelityDiagnostics.Flatten([report]);
        Facts = new ReadOnlyDictionary<string, string>(new Dictionary<string, string>(facts ?? new Dictionary<string, string>(), StringComparer.Ordinal));
    }

    internal OfficeWorkflowConversionEvidence(IReadOnlyList<IOfficeConversionReport> reports, IReadOnlyDictionary<string, string> facts) {
        FidelityDiagnostics = OfficeConversionFidelityDiagnostics.Flatten(reports);
        Facts = new ReadOnlyDictionary<string, string>(new Dictionary<string, string>(facts, StringComparer.Ordinal));
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
