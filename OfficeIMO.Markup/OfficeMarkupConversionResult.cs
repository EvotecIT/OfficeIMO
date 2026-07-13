namespace OfficeIMO.Markup;

/// <summary>Structured failure raised when OfficeIMO Markup cannot be converted safely.</summary>
public sealed class OfficeMarkupConversionException : InvalidOperationException {
    /// <summary>Creates an exception from the diagnostics emitted by a conversion.</summary>
    public OfficeMarkupConversionException(IReadOnlyList<OfficeMarkupDiagnostic> diagnostics)
        : base(BuildMessage(diagnostics)) {
        Diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
    }

    /// <summary>Diagnostics captured by the failed or lossy conversion.</summary>
    public IReadOnlyList<OfficeMarkupDiagnostic> Diagnostics { get; }

    private static string BuildMessage(IReadOnlyList<OfficeMarkupDiagnostic>? diagnostics) {
        OfficeMarkupDiagnostic? first = diagnostics?.FirstOrDefault(static diagnostic =>
            diagnostic.Severity != OfficeMarkupDiagnosticSeverity.Info);
        return first == null
            ? "OfficeIMO Markup conversion failed."
            : "OfficeIMO Markup conversion failed: " + first.Message;
    }
}

/// <summary>Immutable diagnostics from one OfficeIMO Markup conversion.</summary>
public sealed class OfficeMarkupConversionReport {
    private readonly IReadOnlyList<OfficeMarkupDiagnostic> _diagnostics;

    /// <summary>Creates a conversion report.</summary>
    public OfficeMarkupConversionReport(IEnumerable<OfficeMarkupDiagnostic>? diagnostics = null) {
        _diagnostics = Array.AsReadOnly((diagnostics ?? Array.Empty<OfficeMarkupDiagnostic>()).ToArray());
    }

    /// <summary>Immutable conversion diagnostics in emission order.</summary>
    public IReadOnlyList<OfficeMarkupDiagnostic> Diagnostics => _diagnostics;

    /// <summary>Whether conversion completed without an error diagnostic.</summary>
    public bool Succeeded => !_diagnostics.Any(static diagnostic =>
        diagnostic.Severity == OfficeMarkupDiagnosticSeverity.Error);

    /// <summary>Whether conversion warned about, omitted, or failed any source content.</summary>
    public bool HasLoss => _diagnostics.Any(static diagnostic =>
        diagnostic.Severity != OfficeMarkupDiagnosticSeverity.Info);

    /// <summary>Throws when conversion failed.</summary>
    public void RequireSuccess() {
        if (!Succeeded) throw new OfficeMarkupConversionException(_diagnostics);
    }

    /// <summary>Throws when conversion failed or reported possible content loss.</summary>
    public void RequireNoLoss() {
        if (HasLoss) throw new OfficeMarkupConversionException(_diagnostics);
    }
}

/// <summary>A native target document and immutable diagnostics from one OfficeIMO Markup conversion.</summary>
/// <typeparam name="TDocument">Native target document type.</typeparam>
public class OfficeMarkupConversionResult<TDocument> where TDocument : class {
    /// <summary>Creates a conversion result.</summary>
    public OfficeMarkupConversionResult(TDocument value, IEnumerable<OfficeMarkupDiagnostic>? diagnostics = null)
        : this(value, new OfficeMarkupConversionReport(diagnostics)) {
    }

    /// <summary>Creates a conversion result from an existing immutable report.</summary>
    public OfficeMarkupConversionResult(TDocument value, OfficeMarkupConversionReport report) {
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Report = report ?? throw new ArgumentNullException(nameof(report));
    }

    /// <summary>Native target document. The caller owns and disposes it when applicable.</summary>
    public TDocument Value { get; }

    /// <summary>Immutable conversion report.</summary>
    public OfficeMarkupConversionReport Report { get; }

    /// <summary>Immutable conversion diagnostics in emission order.</summary>
    public IReadOnlyList<OfficeMarkupDiagnostic> Diagnostics => Report.Diagnostics;

    /// <summary>Whether conversion completed without an error diagnostic.</summary>
    public bool Succeeded => Report.Succeeded;

    /// <summary>Whether conversion warned about, omitted, or failed any source content.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>Returns the native document when conversion succeeded.</summary>
    public TDocument RequireValue() {
        Report.RequireSuccess();
        return Value;
    }

    /// <summary>Returns the native document only when conversion reported no possible loss.</summary>
    public TDocument RequireNoLoss() {
        Report.RequireNoLoss();
        return Value;
    }
}
