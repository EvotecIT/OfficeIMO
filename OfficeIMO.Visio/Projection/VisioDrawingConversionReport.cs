namespace OfficeIMO.Visio;

/// <summary>Source-qualified fidelity diagnostics for one Visio diagram-page projection.</summary>
public sealed class VisioDrawingConversionReport : IOfficeConversionReport {
    private readonly List<OfficeConversionFidelityDiagnostic> _diagnostics = new();

    internal VisioDrawingConversionReport(string? sourceName) { SourceName = sourceName; }

    /// <summary>Source name associated with this operation.</summary>
    public string? SourceName { get; }

    /// <inheritdoc />
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics => _diagnostics.AsReadOnly();

    /// <inheritdoc />
    public bool HasLoss => _diagnostics.Any(item => item.LossKind != OfficeConversionLossKind.None);

    /// <inheritdoc />
    public void RequireNoLoss() {
        if (HasLoss) throw new OfficeConversionException("Visio diagram-page projection reported fidelity loss.", this);
    }

    internal void Add(string code, string message, OfficeConversionLossKind kind, string location) =>
        _diagnostics.Add(new OfficeConversionFidelityDiagnostic(code, message, kind, "OfficeIMO.Visio.Drawing", location));
}
