namespace OfficeIMO.Excel;

/// <summary>Immutable omissions and approximations for one XML Spreadsheet interchange operation.</summary>
public sealed class ExcelSpreadsheetXmlReport : IOfficeConversionReport {
    internal ExcelSpreadsheetXmlReport(IEnumerable<OfficeConversionFidelityDiagnostic> diagnostics) =>
        FidelityDiagnostics = Array.AsReadOnly(diagnostics.ToArray());

    /// <inheritdoc />
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; }
    /// <inheritdoc />
    public bool HasLoss => FidelityDiagnostics.Any(item => item.LossKind != OfficeConversionLossKind.None);
    /// <inheritdoc />
    public void RequireNoLoss() {
        if (HasLoss) throw new OfficeConversionException("XML Spreadsheet interchange reported content loss. Review FidelityDiagnostics before using the result.", this);
    }
}
