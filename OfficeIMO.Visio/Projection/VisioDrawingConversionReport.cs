namespace OfficeIMO.Visio;

/// <summary>Source-qualified fidelity diagnostics for one Visio diagram-page projection.</summary>
public sealed class VisioDrawingConversionReport : OfficeConversionReport {
    private readonly List<OfficeDiagnostic> _diagnostics = new();

    internal VisioDrawingConversionReport(string? sourceName) { SourceName = sourceName; }

    /// <summary>Source name associated with this operation.</summary>
    public string? SourceName { get; }

    /// <inheritdoc />
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics => _diagnostics.AsReadOnly();

    /// <inheritdoc />
    public override bool HasLoss => _diagnostics.Any(item => item.LossKind != OfficeConversionLossKind.None);

    /// <inheritdoc />
    protected override OfficeConversionException CreateLossException() => new OfficeConversionException("Visio diagram-page projection reported fidelity loss.", this);

    internal void Add(string code, string message, OfficeConversionLossKind kind, string location) =>
        _diagnostics.Add(new global::OfficeIMO.OfficeDiagnostic(code, kind == global::OfficeIMO.OfficeConversionLossKind.None ? global::OfficeIMO.OfficeDiagnosticSeverity.Information : global::OfficeIMO.OfficeDiagnosticSeverity.Warning, message, location, "OfficeIMO.Visio.Drawing", kind));
}
