namespace OfficeIMO.Excel;

/// <summary>Immutable omissions and approximations for one XML Spreadsheet interchange operation.</summary>
public sealed class ExcelSpreadsheetXmlReport : OfficeConversionReport {
    internal ExcelSpreadsheetXmlReport(IEnumerable<OfficeDiagnostic> diagnostics) =>
        FidelityDiagnostics = Array.AsReadOnly(diagnostics.ToArray());

    /// <inheritdoc />
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics { get; }
    /// <inheritdoc />
    public override bool HasLoss => FidelityDiagnostics.Any(item => item.LossKind != OfficeConversionLossKind.None);
    /// <inheritdoc />
    protected override OfficeConversionException CreateLossException() => new OfficeConversionException("XML Spreadsheet interchange reported content loss. Review FidelityDiagnostics before using the result.", this);
}
