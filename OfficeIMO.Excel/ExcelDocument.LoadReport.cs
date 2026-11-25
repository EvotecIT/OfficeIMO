using System.Diagnostics;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    private OfficeDiagnostic[] _sourceLoadDiagnostics = Array.Empty<OfficeDiagnostic>();

    /// <summary>Loads a workbook and returns findings from its source projection.</summary>
    /// <remarks>
    /// Worksheet cells load on first use, and the report includes findings for each worksheet as it is loaded, so read
    /// <see cref="OfficeConversionReport.Diagnostics"/> after reading the cells. For a complete check up front, call
    /// <see cref="ValidateDocument"/>, which parses every worksheet into the Open XML SDK model and reports schema errors
    /// such as an empty cell style attribute.
    /// </remarks>
    public static ExcelDocument Load(string path, ExcelLoadOptions? options, out OfficeConversionReport report) {
        var timer = Stopwatch.StartNew();
        ExcelDocument value = Load(path, options);
        report = value.CreateLoadReport(path, timer.Elapsed);
        return value;
    }
    /// <summary>Loads a workbook stream and returns findings from its source projection.</summary>
    /// <remarks>
    /// Worksheet cells load on first use, and the report includes findings for each worksheet as it is loaded, so read
    /// <see cref="OfficeConversionReport.Diagnostics"/> after reading the cells. For a complete check up front, call
    /// <see cref="ValidateDocument"/>, which parses every worksheet into the Open XML SDK model and reports schema errors
    /// such as an empty cell style attribute.
    /// </remarks>
    public static ExcelDocument Load(Stream stream, ExcelLoadOptions? options, out OfficeConversionReport report) {
        var timer = Stopwatch.StartNew();
        ExcelDocument value = Load(stream, options);
        report = value.CreateLoadReport("Excel stream", timer.Elapsed);
        return value;
    }
    internal OfficeConversionReport CreateLoadReport(string source, TimeSpan elapsed) =>
        new ExcelLoadReport(this, source, Sheets.Count, elapsed);

    /// <summary>Adds a finding about the source that was found after loading, such as a repair made while reading a worksheet.</summary>
    internal void AddSourceLoadDiagnostic(OfficeDiagnostic diagnostic) {
        while (true) {
            OfficeDiagnostic[] current = Volatile.Read(ref _sourceLoadDiagnostics);
            var findings = new OfficeDiagnostic[current.Length + 1];
            Array.Copy(current, findings, current.Length);
            findings[findings.Length - 1] = diagnostic;
            if (Interlocked.CompareExchange(ref _sourceLoadDiagnostics, findings, current) == current) return;
        }
    }

    /// <summary>Reads the document's source findings when they are read, including repairs recorded as worksheets load.</summary>
    private sealed class ExcelLoadReport : OfficeConversionReport {
        private static readonly IReadOnlyList<OfficeDiagnostic> NoFindings =
            new System.Collections.ObjectModel.ReadOnlyCollection<OfficeDiagnostic>(Array.Empty<OfficeDiagnostic>());
        private readonly ExcelDocument _document;
        private FindingsView? _view;

        internal ExcelLoadReport(ExcelDocument document, string source, int itemsWritten, TimeSpan elapsed)
            : base(source, "Excel", itemsWritten, elapsed) {
            _document = document;
        }

        // Callers get a read-only view; it is rebuilt only after the document records another finding.
        protected override IReadOnlyList<OfficeDiagnostic> GetDiagnostics() {
            OfficeDiagnostic[] findings = Volatile.Read(ref _document._sourceLoadDiagnostics);
            if (findings.Length == 0) return NoFindings;
            FindingsView? view = Volatile.Read(ref _view);
            if (view == null || !ReferenceEquals(view.Source, findings)) {
                view = new FindingsView(findings);
                Volatile.Write(ref _view, view);
            }
            return view.Items;
        }

        private sealed class FindingsView {
            internal FindingsView(OfficeDiagnostic[] source) {
                Source = source;
                Items = Array.AsReadOnly(source);
            }

            internal OfficeDiagnostic[] Source { get; }
            internal IReadOnlyList<OfficeDiagnostic> Items { get; }
        }
    }

    private void CaptureLegacyLoadDiagnostics(LegacyXls.Model.LegacyXlsWorkbook workbook) {
        var findings = new LegacyXlsLoadResult(this, workbook).Diagnostics;
        _sourceLoadDiagnostics = findings.Select(d => new OfficeDiagnostic(d.Code,
            workbook.Diagnostics.FirstOrDefault(item => item.Code == d.Code)?.Severity switch {
                LegacyXlsDiagnosticSeverity.Info => OfficeDiagnosticSeverity.Information,
                LegacyXlsDiagnosticSeverity.Error => OfficeDiagnosticSeverity.Error,
                _ => OfficeDiagnosticSeverity.Warning
            }, d.Message, d.Location, d.Format, d.LossKind)).ToArray();
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
