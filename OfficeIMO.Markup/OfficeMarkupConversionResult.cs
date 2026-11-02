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
public sealed class OfficeMarkupConversionReport : OfficeConversionReport {
    private readonly IReadOnlyList<OfficeMarkupDiagnostic> _diagnostics;

    /// <summary>Creates a conversion report.</summary>
    public OfficeMarkupConversionReport(IEnumerable<OfficeMarkupDiagnostic>? diagnostics = null) {
        _diagnostics = Array.AsReadOnly((diagnostics ?? Array.Empty<OfficeMarkupDiagnostic>()).ToArray());
        FidelityDiagnostics = Array.AsReadOnly(_diagnostics.Select(static (diagnostic, index) =>
            new global::OfficeIMO.OfficeDiagnostic(
                "OFFICE_MARKUP_" + diagnostic.Severity.ToString().ToUpperInvariant(),
                diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.None ? global::OfficeIMO.OfficeDiagnosticSeverity.Information : global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
                string.IsNullOrWhiteSpace(diagnostic.Message) ? "Office markup diagnostic." : diagnostic.Message,
                diagnostic.Node == null ? index.ToString(System.Globalization.CultureInfo.InvariantCulture) : diagnostic.Node.Kind.ToString(),
                "OfficeIMO.Markup",
                diagnostic.LossKind)).ToArray());
    }

    /// <summary>Immutable conversion diagnostics in emission order.</summary>
    public IReadOnlyList<OfficeMarkupDiagnostic> NativeDiagnostics => _diagnostics;

    /// <summary>Category-preserving conversion diagnostics.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics { get; }

    /// <summary>Whether conversion completed without an error diagnostic.</summary>
    public bool Succeeded => !_diagnostics.Any(static diagnostic =>
        diagnostic.Severity == OfficeMarkupDiagnosticSeverity.Error);

    /// <summary>Whether conversion approximated, omitted, or failed any source content.</summary>
    public override bool HasLoss => _diagnostics.Any(static diagnostic => diagnostic.LossKind != OfficeConversionLossKind.None);

    /// <summary>Throws when conversion failed.</summary>
    public void RequireSuccess() {
        if (!Succeeded) throw new OfficeMarkupConversionException(_diagnostics);
    }

}

/// <summary>A native target document and immutable diagnostics from one OfficeIMO Markup conversion.</summary>
/// <typeparam name="TDocument">Native target document type.</typeparam>
public class OfficeMarkupConversionResult<TDocument> : OfficeConversionResult<TDocument> where TDocument : class {
    /// <summary>Gets the typed conversion report.</summary>
    public new OfficeMarkupConversionReport Report => (OfficeMarkupConversionReport)base.Report;

    /// <summary>Creates a conversion result.</summary>
    public OfficeMarkupConversionResult(TDocument value, IEnumerable<OfficeMarkupDiagnostic>? diagnostics = null)
        : this(value, new OfficeMarkupConversionReport(diagnostics)) {
    }

    /// <summary>Creates a conversion result from an existing immutable report.</summary>
    public OfficeMarkupConversionResult(TDocument value, OfficeMarkupConversionReport report)
        : base(value, report) { }

    /// <summary>Immutable conversion diagnostics in emission order.</summary>
    public IReadOnlyList<OfficeMarkupDiagnostic> Diagnostics => Report.NativeDiagnostics;

    /// <summary>Whether conversion completed without an error diagnostic.</summary>
    public bool Succeeded => Report.Succeeded;

    /// <summary>Returns the native document or throws the markup-specific exception when conversion failed.</summary>
    public TDocument RequireValue() {
        Report.RequireSuccess();
        return Value;
    }
}
