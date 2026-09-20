namespace OfficeIMO.Bibliography;

/// <summary>How a source feature was handled at a bibliography conversion boundary.</summary>
public enum BibliographyConversionAction {
    /// <summary>Mapped without known semantic loss.</summary>
    Mapped = 0,
    /// <summary>Retained as a destination-native extension.</summary>
    PreservedExtension,
    /// <summary>Represented with reduced precision or changed shape.</summary>
    Approximated,
    /// <summary>Not written to the destination.</summary>
    Omitted
}

/// <summary>One deterministic conversion decision.</summary>
public sealed class BibliographyConversionDiagnostic {
    /// <summary>Initializes a conversion diagnostic.</summary>
    public BibliographyConversionDiagnostic(string code, BibliographyDiagnosticSeverity severity, string message, BibliographyConversionAction action, string? itemKey = null, string? field = null) {
        Code = code ?? throw new ArgumentNullException(nameof(code));
        Severity = severity;
        Message = message ?? throw new ArgumentNullException(nameof(message));
        Action = action;
        ItemKey = itemKey;
        Field = field;
    }

    /// <summary>Stable machine-readable code.</summary>
    public string Code { get; }
    /// <summary>Severity.</summary>
    public BibliographyDiagnosticSeverity Severity { get; }
    /// <summary>Description.</summary>
    public string Message { get; }
    /// <summary>Conversion action.</summary>
    public BibliographyConversionAction Action { get; }
    /// <summary>Related citation key.</summary>
    public string? ItemKey { get; }
    /// <summary>Related field.</summary>
    public string? Field { get; }
}

/// <summary>Fidelity evidence for one bibliography write or conversion.</summary>
public sealed class BibliographyConversionReport : global::OfficeIMO.IOfficeConversionReport {
    private readonly List<BibliographyConversionDiagnostic> _diagnostics = new List<BibliographyConversionDiagnostic>();

    /// <summary>Conversion decisions in deterministic item and field order.</summary>
    public IReadOnlyList<BibliographyConversionDiagnostic> Diagnostics => _diagnostics.AsReadOnly();
    /// <summary>Category-preserving diagnostics for composed conversion routes.</summary>
    public IReadOnlyList<global::OfficeIMO.OfficeConversionFidelityDiagnostic> FidelityDiagnostics =>
        Array.AsReadOnly(_diagnostics.Select(static diagnostic => new global::OfficeIMO.OfficeConversionFidelityDiagnostic(
            diagnostic.Code,
            diagnostic.Message,
            GetLossKind(diagnostic),
            "OfficeIMO.Bibliography",
            GetLocation(diagnostic))).ToArray());
    /// <summary>True when any value was approximated, omitted, or failed.</summary>
    public bool HasLoss => _diagnostics.Any(static diagnostic => GetLossKind(diagnostic) != global::OfficeIMO.OfficeConversionLossKind.None);

    internal void Add(string code, BibliographyDiagnosticSeverity severity, string message, BibliographyConversionAction action, BibliographyItem? item = null, string? field = null) =>
        _diagnostics.Add(new BibliographyConversionDiagnostic(code, severity, message, action, item?.Key, field));

    /// <summary>Adds an adapter or caller conversion decision.</summary>
    public void Add(BibliographyConversionDiagnostic diagnostic) => _diagnostics.Add(diagnostic ?? throw new ArgumentNullException(nameof(diagnostic)));

    /// <summary>Adds decisions from another report in their original order.</summary>
    public void Merge(BibliographyConversionReport? report) {
        if (report == null || ReferenceEquals(report, this)) return;
        foreach (BibliographyConversionDiagnostic diagnostic in report.Diagnostics) _diagnostics.Add(diagnostic);
    }

    /// <summary>Throws when the report contains conversion loss.</summary>
    public void RequireNoLoss() {
        if (HasLoss) throw new BibliographyConversionLossException(this);
    }

    private static global::OfficeIMO.OfficeConversionLossKind GetLossKind(BibliographyConversionDiagnostic diagnostic) {
        if (diagnostic.Severity == BibliographyDiagnosticSeverity.Error) return global::OfficeIMO.OfficeConversionLossKind.Failure;
        return diagnostic.Action switch {
            BibliographyConversionAction.Approximated => global::OfficeIMO.OfficeConversionLossKind.Approximation,
            BibliographyConversionAction.Omitted => global::OfficeIMO.OfficeConversionLossKind.Omission,
            _ => global::OfficeIMO.OfficeConversionLossKind.None
        };
    }

    private static string? GetLocation(BibliographyConversionDiagnostic diagnostic) {
        if (string.IsNullOrWhiteSpace(diagnostic.ItemKey)) return diagnostic.Field;
        return string.IsNullOrWhiteSpace(diagnostic.Field)
            ? diagnostic.ItemKey
            : diagnostic.ItemKey + ":" + diagnostic.Field;
    }
}

/// <summary>Thrown when strict conversion rejects a lossy result.</summary>
public sealed class BibliographyConversionLossException : InvalidOperationException {
    internal BibliographyConversionLossException(BibliographyConversionReport report)
        : base("The bibliography conversion would approximate or omit source data. Inspect Report.Diagnostics for details.") {
        Report = report;
    }

    /// <summary>Conversion report that caused the failure.</summary>
    public BibliographyConversionReport Report { get; }
}
