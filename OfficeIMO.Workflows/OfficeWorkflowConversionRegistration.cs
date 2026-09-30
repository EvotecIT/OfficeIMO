using System.Collections.ObjectModel;

namespace OfficeIMO.Workflows;

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
        Func<IOfficeWorkflowConversionSettings, bool>? acceptsSettings) {
        OfficeWorkflowRoute route = OfficeWorkflowCatalog.Find(routeId)
            ?? throw new ArgumentException("Choose an existing canonical conversion route.", nameof(routeId));
        if (route.CanExecute) throw new ArgumentException("Built-in conversion routes cannot be replaced.", nameof(routeId));
        if (route.TargetExtension.TrimStart('.').ToLowerInvariant() is not ("docx" or "xlsx" or "pptx"))
            throw new NotSupportedException("Opt-in workflow conversion currently supports DOCX, XLSX, and PPTX destinations.");
        RouteId = route.Id;
        Converter = converter;
        _acceptsSettings = acceptsSettings;
    }

    /// <summary>Canonical capability identifier.</summary>
    public string RouteId { get; }
    internal Func<Stream, Stream, OfficeWorkflowLimits, IOfficeWorkflowConversionSettings?, CancellationToken, OfficeWorkflowConversionEvidence> Converter { get; }
    private readonly Func<IOfficeWorkflowConversionSettings, bool>? _acceptsSettings;

    /// <summary>Registers a canonical route with an explicit adapter-owned settings contract.</summary>
    public static OfficeWorkflowConversionRegistration Create<TSettings>(string routeId,
        OfficeWorkflowConfiguredConverter<TSettings> converter) where TSettings : class, IOfficeWorkflowConversionSettings {
        ArgumentNullException.ThrowIfNull(converter);
        return new(routeId, (input, output, limits, settings, token) => converter(input, output, limits, (TSettings?)settings, token),
            settings => settings is TSettings);
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
