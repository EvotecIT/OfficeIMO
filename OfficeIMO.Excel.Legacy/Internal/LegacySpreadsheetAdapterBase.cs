using System;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Excel.Legacy;

internal abstract class LegacySpreadsheetAdapterBase : ILegacySpreadsheetAdapter {
    public abstract LegacySpreadsheetFormat Format { get; }
    public abstract string ProfileId { get; }
    public virtual string GetProfileId(byte[] data) => ProfileId;
    public abstract int Probe(byte[] data, string? sourceName, out string reason);
    public abstract LegacySpreadsheetModel Parse(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellationToken);

    protected static bool ExtensionIs(string? sourceName, params string[] extensions) {
        string extension = Path.GetExtension(sourceName ?? string.Empty);
        return extensions.Any(candidate => string.Equals(extension, candidate, StringComparison.OrdinalIgnoreCase));
    }

    protected static OfficeCompatibilityFinding Loss(string code, string category, string message) =>
        new(code, category, message, OfficeCompatibilityState.Approximated, OfficeCompatibilitySeverity.Warning,
            OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Visual | OfficeCompatibilityImpact.Carrier, true);

    protected static OfficeCompatibilityFinding Inert(string code, string category, string message) =>
        new(code, category, message, OfficeCompatibilityState.Dropped, OfficeCompatibilitySeverity.Warning,
            OfficeCompatibilityImpact.Behavioral | OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier, true);

    internal static OfficeCompatibilityFinding LossFinding(string code, string category, string message) => Loss(code, category, message);

    protected static LegacySpreadsheetModel ParseDelimitedSalvage(byte[] data, OfficeLegacyImportLimits limits, string limitation, CancellationToken cancellationToken) {
        string text = OfficeLegacyImportBuffer.ExtractPrintableText(data, 0, data.Length, limits.MaxTextCharacters, false, cancellationToken: cancellationToken);
        if (string.IsNullOrWhiteSpace(text)) throw new InvalidDataException("Legacy spreadsheet did not contain recoverable bounded text.");
        var model = new LegacySpreadsheetModel { Quality = OfficeLegacyImportQuality.Salvage };
        var sheet = new LegacySpreadsheetSheet("Sheet1");
        model.Sheets.Add(sheet);
        int row = 1;
        foreach (string line in text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')) {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(line)) continue;
            if (row > 1048576) throw new InvalidDataException("Legacy spreadsheet row is outside the supported workbook model.");
            char separator = line.IndexOf('\t') >= 0 ? '\t' : line.IndexOf(',') >= 0 ? ',' : '\0';
            string[] fields = separator == '\0' ? new[] { line } : line.Split(separator);
            for (int column = 0; column < fields.Length; column++) {
                if (column >= 16384) throw new InvalidDataException("Legacy spreadsheet column is outside the supported workbook model.");
                if (model.RecoveredCellCount >= limits.MaxItems) throw new InvalidDataException("Legacy spreadsheet exceeds the configured cell limit.");
                sheet.Cells.Add(new LegacySpreadsheetCell(row, column + 1, fields[column]));
                model.RecoveredCellCount++;
            }
            row++;
        }
        model.Findings.Add(Loss("LEGACY_SHEET_SALVAGE", "Structure", limitation));
        return model;
    }
}
