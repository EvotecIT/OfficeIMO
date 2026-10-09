using System.Globalization;
using System.Threading;

namespace OfficeIMO.Excel.Legacy;

/// <summary>Imports a single SYLK worksheet as stored values; source expressions are never evaluated.</summary>
internal sealed class SylkAdapter : TextSpreadsheetAdapterBase {
    public override LegacySpreadsheetFormat Format => LegacySpreadsheetFormat.Sylk;
    public override string ProfileId => "sylk-stored-values";

    public override int Probe(byte[] data, string? sourceName, OfficeLegacyImportLimits limits, CancellationToken cancellationToken, out string reason) {
        cancellationToken.ThrowIfCancellationRequested();
        bool matches = TextSpreadsheetReader.HasHeader(data, "ID;");
        reason = matches ? "SYLK ID record signature." : "No SYLK ID record.";
        return matches ? 95 : 0;
    }

    protected override LegacySpreadsheetModel Parse(TextSpreadsheetReader reader) {
        string header = reader.RequiredLine();
        if (!header.StartsWith("ID;", StringComparison.Ordinal)) throw new InvalidDataException("SYLK must start with an ID record.");
        bool modernCalc = header.StartsWith("ID;PCALCOOO32", StringComparison.Ordinal);
        var model = new LegacySpreadsheetModel { Quality = OfficeLegacyImportQuality.Structured };
        var sheet = new LegacySpreadsheetSheet("Sheet1");
        model.Sheets.Add(sheet);
        var addresses = new HashSet<long>();
        var omittedRecords = new HashSet<string>(StringComparer.Ordinal);
        int row = 1, column = 1, formulaCount = 0, missingCacheCount = 0, invalidCacheCount = 0, errorCount = 0;
        bool formatting = false;
        string? line;
        while ((line = reader.ReadLine()) != null) {
            if (line.Length == 0) continue;
            if (line == "E") {
                reader.RequireEnd();
                if (formulaCount > 0) {
                    model.Metadata["OmittedFormulaCount"] = formulaCount.ToString(CultureInfo.InvariantCulture);
                    model.Findings.Add(Loss("SYLK_FORMULA_STORED_VALUE", "Formula", "SYLK formulas were omitted without evaluation or link resolution; only valid stored values were imported."));
                }
                if (missingCacheCount > 0) {
                    model.Metadata["MissingFormulaValueCount"] = missingCacheCount.ToString(CultureInfo.InvariantCulture);
                    model.Findings.Add(Loss("SYLK_FORMULA_VALUE_MISSING", "Formula", "One or more source formulas had no valid stored value and were imported as blank cells."));
                }
                if (invalidCacheCount > 0) model.Findings.Add(Loss("SYLK_STORED_VALUE_INVALID", "Cell", "Source values marked invalid by SYLK were imported as blank cells."));
                if (errorCount > 0) model.Findings.Add(Loss("SYLK_ERROR_AS_TEXT", "Cell", "SYLK error values were retained as literal text rather than live Excel error cells."));
                if (formatting) model.Findings.Add(Loss("SYLK_FORMATTING_OMITTED", "Formatting", "SYLK format definitions, number formats, widths and style records were omitted; numeric dates remain stored numbers."));
                if (omittedRecords.Count > 0) model.Findings.Add(Loss("SYLK_RECORDS_OMITTED", "Structure", "Unsupported SYLK records or cell attributes were omitted: " + string.Join(", ", omittedRecords.OrderBy(static item => item, StringComparer.Ordinal)) + "."));
                return model;
            }
            List<string> fields = SplitFields(line);
            string record = fields[0];
            if (record == "P" || record == "F") {
                formatting = true;
                if (record == "F") {
                    foreach (string field in fields.Skip(1)) {
                        if (field.StartsWith("X", StringComparison.Ordinal)) column = TextSpreadsheetReader.Coordinate(field.Substring(1), 16_384);
                        if (field.StartsWith("Y", StringComparison.Ordinal)) row = TextSpreadsheetReader.Coordinate(field.Substring(1), 1_048_576);
                    }
                }
                continue;
            }
            if (record == "B") continue; // Source dimensions do not cause allocation or synthesize cells.
            if (record != "C") { AddOmitted(omittedRecords, record); continue; }
            object? value = null;
            bool hasValue = false, hasFormula = false, invalidValue = false;
            foreach (string field in fields.Skip(1)) {
                if (field.Length == 0) throw new InvalidDataException("Empty SYLK cell field.");
                string content = field.Substring(1);
                switch (field[0]) {
                    case 'X': column = TextSpreadsheetReader.Coordinate(content, 16_384); break;
                    case 'Y': row = TextSpreadsheetReader.Coordinate(content, 1_048_576); break;
                    case 'K':
                        if (hasValue) throw new InvalidDataException("Duplicate SYLK stored-value field.");
                        hasValue = true;
                        value = ParseValue(content, modernCalc);
                        if (value is string && !content.StartsWith("\"", StringComparison.Ordinal)) errorCount++;
                        break;
                    case 'E': case 'M': case 'S': hasFormula = true; break;
                    case 'I': invalidValue = true; break;
                    case 'R': case 'C': hasFormula = true; break; // Formula reference coordinates, never resolved.
                    default: AddOmitted(omittedRecords, "C;" + field[0]); break;
                }
            }
            if (hasFormula) formulaCount++;
            if (invalidValue) { value = null; hasValue = false; }
            if (hasFormula && !hasValue) missingCacheCount++;
            if (!addresses.Add(((long)row << 16) | (uint)column)) throw new InvalidDataException("SYLK defines the same cell more than once.");
            if (invalidValue && !hasFormula) invalidCacheCount++;
            reader.AddCell(model, sheet, row, column, value);
        }
        throw new InvalidDataException("SYLK is missing its E terminator.");
    }

    private static object ParseValue(string value, bool modernCalc) {
        if (value.StartsWith("\"", StringComparison.Ordinal)) {
            if (value.Length < 2 || !value.EndsWith("\"", StringComparison.Ordinal)) throw new InvalidDataException("Unterminated SYLK text value.");
            string text = value.Substring(1, value.Length - 2);
            if (!modernCalc) text = text.Replace("\"\"", "\"");
            text = text.Replace("\u001b :", "\n");
            if (text.IndexOf('\u001b') >= 0) throw new NotSupportedException("SYLK character escape sequences other than the line-break escape are not supported; choose an explicit text encoding for unescaped input.");
            return text;
        }
        if (value == "TRUE") return true;
        if (value == "FALSE") return false;
        if (value == "#NULL!" || value == "#DIV/0!" || value == "#VALUE!" || value == "#REF!" || value == "#NAME?" || value == "#NUM!" || value == "#N/A") return value;
        return TextSpreadsheetReader.Number(value);
    }

    private static List<string> SplitFields(string line) {
        var fields = new List<string>();
        var field = new System.Text.StringBuilder();
        for (int index = 0; index < line.Length; index++) {
            if (line[index] == ';') {
                if (index + 1 < line.Length && line[index + 1] == ';') { field.Append(';'); index++; }
                else {
                    if (fields.Count >= 64) throw new InvalidDataException("SYLK record has more than 64 fields.");
                    fields.Add(field.ToString()); field.Clear();
                }
            } else field.Append(line[index]);
        }
        if (fields.Count >= 64) throw new InvalidDataException("SYLK record has more than 64 fields.");
        fields.Add(field.ToString());
        return fields;
    }

    private static void AddOmitted(HashSet<string> records, string record) {
        if (records.Count < 64) records.Add(record.Length > 16 ? record.Substring(0, 16) : record);
        else records.Add("Additional record kinds");
    }
}
