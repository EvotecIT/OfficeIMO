using System;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Excel.Legacy;

internal abstract class LegacySpreadsheetAdapterBase : ILegacySpreadsheetAdapter {
    protected const int ExcelCellTextLimit = 32_767;
    public abstract LegacySpreadsheetFormat Format { get; }
    public abstract string ProfileId { get; }
    public virtual string GetProfileId(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return ProfileId;
    }
    public abstract int Probe(byte[] data, string? sourceName, OfficeLegacyImportLimits limits, CancellationToken cancellationToken, out string reason);
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
        int inspectedRecords = 0;
        int truncatedCellCount = 0;
        string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        for (int lineStart = 0; lineStart <= normalized.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            if (++inspectedRecords > limits.MaxRecords) throw new InvalidDataException("Legacy spreadsheet exceeds the configured record limit.");
            int lineEnd = normalized.IndexOf('\n', lineStart);
            if (lineEnd < 0) lineEnd = normalized.Length;
            int lineLength = lineEnd - lineStart;
            if (!IsWhiteSpace(normalized, lineStart, lineLength)) {
                if (row > 1048576) throw new InvalidDataException("Legacy spreadsheet row is outside the supported workbook model.");
                AddSalvageRow(normalized, lineStart, lineLength, row, sheet, model, limits, ref truncatedCellCount);
                row++;
            }
            if (lineEnd == normalized.Length) break;
            lineStart = lineEnd + 1;
        }
        if (truncatedCellCount > 0) {
            model.Metadata["TruncatedSalvageCellCount"] = truncatedCellCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
            AddCellTextTruncationFinding(model);
        }
        model.Findings.Add(Loss("LEGACY_SHEET_SALVAGE", "Structure", limitation));
        return model;
    }

    private static void AddSalvageRow(
        string text,
        int lineStart,
        int lineLength,
        int row,
        LegacySpreadsheetSheet sheet,
        LegacySpreadsheetModel model,
        OfficeLegacyImportLimits limits,
        ref int truncatedCellCount) {
        int lineEnd = lineStart + lineLength;
        int tab = text.IndexOf('\t', lineStart, lineLength);
        int comma = tab < 0 ? text.IndexOf(',', lineStart, lineLength) : -1;
        char separator = tab >= 0 ? '\t' : comma >= 0 ? ',' : '\0';
        int fieldStart = lineStart;
        for (int column = 1; ; column++) {
            if (column > 16384) throw new InvalidDataException("Legacy spreadsheet column is outside the supported workbook model.");
            if (model.RecoveredCellCount >= limits.MaxItems) throw new InvalidDataException("Legacy spreadsheet exceeds the configured cell limit.");

            int separatorIndex = separator == '\0' ? -1 : text.IndexOf(separator, fieldStart, lineEnd - fieldStart);
            int fieldEnd = separatorIndex < 0 ? lineEnd : separatorIndex;
            int fieldLength = fieldEnd - fieldStart;
            int retainedLength = Math.Min(fieldLength, ExcelCellTextLimit);
            string value = text.Substring(fieldStart, retainedLength);
            if (fieldLength > retainedLength) truncatedCellCount++;
            sheet.Cells.Add(new LegacySpreadsheetCell(row, column, value));
            model.RecoveredCellCount++;

            if (separatorIndex < 0) break;
            fieldStart = separatorIndex + 1;
        }
    }

    private static bool IsWhiteSpace(string text, int start, int length) {
        int end = start + length;
        for (int index = start; index < end; index++) {
            if (!char.IsWhiteSpace(text[index])) return false;
        }
        return true;
    }

    protected static void AddCellTextTruncationFinding(LegacySpreadsheetModel model) =>
        model.Findings.Add(Loss("LEGACY_SHEET_CELL_TEXT_TRUNCATED", "Cell", "One or more recovered text cells exceeded the Excel cell-text limit and were truncated to 32,767 characters; the total is available in metadata."));
}
