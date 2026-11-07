using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class ExcelHtmlConverterExtensions {
    private static string FormatDiagnosticMessage(string message, int occurrenceCount) => occurrenceCount > 1
        ? message + " (" + occurrenceCount.ToString(CultureInfo.InvariantCulture) + " occurrences)"
        : message;

    /// <summary>Groups repeated cell losses as they occur, retaining the first location and deterministic report order per sheet.</summary>
    private sealed class CellFidelityDiagnostics {
        private readonly string _sheetName;
        private readonly Dictionary<(string Code, string Message, OfficeConversionLossKind LossKind), int> _indices = new();
        private readonly List<(string Code, string Message, string Source, OfficeConversionLossKind LossKind, int Count)> _entries = new();

        internal CellFidelityDiagnostics(string sheetName) => _sheetName = sheetName;

        internal void Add(string code, string message, int row, int column, string feature, OfficeConversionLossKind lossKind) {
            var key = (code, message, lossKind);
            if (_indices.TryGetValue(key, out int index)) {
                var entry = _entries[index];
                entry.Count++;
                _entries[index] = entry;
                return;
            }
            _indices.Add(key, _entries.Count);
            _entries.Add((code, message, "excel:cell:" + _sheetName + "!" + ExcelA1.CellReference(row, column) + ":" + feature, lossKind, 1));
        }

        internal void AppendTo(IList<HtmlDiagnostic> diagnostics) {
            foreach (var entry in _entries) {
                diagnostics.Add(new HtmlDiagnostic("OfficeIMO.Excel.Html", entry.Code,
                    FormatDiagnosticMessage(entry.Message, entry.Count), HtmlDiagnosticSeverity.Warning,
                    entry.Source, lossKind: entry.LossKind));
            }
        }
    }

    private static HtmlDiagnostic ToHtmlDiagnostic(OfficeImageExportDiagnostic diagnostic, int occurrenceCount) {
        HtmlDiagnosticSeverity severity = diagnostic.Severity switch {
            OfficeImageExportDiagnosticSeverity.Error => HtmlDiagnosticSeverity.Error,
            OfficeImageExportDiagnosticSeverity.Warning => HtmlDiagnosticSeverity.Warning,
            _ => HtmlDiagnosticSeverity.Info
        };
        return new HtmlDiagnostic(
            "OfficeIMO.Excel.Html",
            diagnostic.Code,
            FormatDiagnosticMessage(diagnostic.Message, occurrenceCount),
            severity,
            diagnostic.Source,
            lossKind: diagnostic.LossKind);
    }

    private static string CreateDiagnosticGroupKey(OfficeImageExportDiagnostic diagnostic) {
        return diagnostic.Severity.ToString() + "|" + diagnostic.Code + "|" + diagnostic.Message;
    }

}
