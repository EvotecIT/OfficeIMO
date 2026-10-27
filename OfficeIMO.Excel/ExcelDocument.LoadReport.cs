using System.Diagnostics;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    private OfficeDiagnostic[] _sourceLoadDiagnostics = Array.Empty<OfficeDiagnostic>();

    /// <summary>Loads a workbook and returns findings from its source projection.</summary>
    public static ExcelDocument Load(string path, ExcelLoadOptions? options, out OfficeConversionReport report) {
        var timer = Stopwatch.StartNew();
        ExcelDocument value = Load(path, options);
        report = value.CreateLoadReport(path, timer.Elapsed);
        return value;
    }
    /// <summary>Loads a workbook stream and returns findings from its source projection.</summary>
    public static ExcelDocument Load(Stream stream, ExcelLoadOptions? options, out OfficeConversionReport report) {
        var timer = Stopwatch.StartNew();
        ExcelDocument value = Load(stream, options);
        report = value.CreateLoadReport("Excel stream", timer.Elapsed);
        return value;
    }
    internal OfficeConversionReport CreateLoadReport(string source, TimeSpan elapsed) =>
        new OfficeConversionReport(source, "Excel", itemsWritten: Sheets.Count, elapsed: elapsed, diagnostics: _sourceLoadDiagnostics);

    private void CaptureLegacyLoadDiagnostics(LegacyXls.Model.LegacyXlsWorkbook workbook) {
        var findings = new LegacyXlsLoadResult(this, workbook).FidelityDiagnostics;
        _sourceLoadDiagnostics = findings.Select(d => new OfficeDiagnostic(d.Code,
            workbook.Diagnostics.FirstOrDefault(item => item.Code == d.Code)?.Severity switch {
                LegacyXlsDiagnosticSeverity.Info => OfficeDiagnosticSeverity.Information,
                LegacyXlsDiagnosticSeverity.Error => OfficeDiagnosticSeverity.Error,
                _ => OfficeDiagnosticSeverity.Warning
            }, d.Message, d.Location, d.Source, d.LossKind)).ToArray();
    }
    private void CaptureBinaryLoadDiagnostics(XlsbWorkbook workbook) {
        var findings = workbook.Diagnostics.Select(d => new OfficeDiagnostic(d.Code,
            d.Severity switch {
                XlsbImportDiagnosticSeverity.Information => OfficeDiagnosticSeverity.Information,
                XlsbImportDiagnosticSeverity.Error => OfficeDiagnosticSeverity.Error,
                _ => OfficeDiagnosticSeverity.Warning
            }, d.Message, d.PartName, "Excel",
            d.Severity == XlsbImportDiagnosticSeverity.Information ? OfficeConversionLossKind.None :
                d.Severity == XlsbImportDiagnosticSeverity.Error ? OfficeConversionLossKind.Failure : OfficeConversionLossKind.Approximation)).ToList();
        foreach (var record in workbook.PreservedRecords)
            findings.Add(new OfficeDiagnostic(OfficeDiagnosticCodes.XlsbUnprojectedRecord, OfficeDiagnosticSeverity.Warning,
                "A preserved binary record is not projected into the editable workbook: " + record.RecordType,
                format: "Excel", lossKind: OfficeConversionLossKind.Omission));
        _sourceLoadDiagnostics = findings.ToArray();
    }
}
