using System.Text.Json;
using OfficeIMO.CSV;

namespace OfficeIMO.Notebook.Spreadsheet;

internal sealed class NotebookTable {
    internal const string Mime = "application/vnd.dataresource+json";
    private const int MaxColumns = 1_024;
    private const int MaxRows = 250_000;
    private const int MaxCells = 1_000_000;
    private const double MinExcelMagnitude = 2.2250738585072014e-308;
    private const double MaxExcelMagnitude = 9.99999999999999e307;

    private readonly IReadOnlyDictionary<string, int> _columns;

    private NotebookTable(string[] headers, Dictionary<string, int> columns, JsonElement rows) {
        Headers = headers;
        _columns = columns;
        Rows = rows;
    }

    internal string[] Headers { get; }
    internal JsonElement Rows { get; }
    internal int RowCount => Rows.GetArrayLength();

    internal static NotebookTable Read(NotebookOutput output, CancellationToken token) {
        ArgumentNullException.ThrowIfNull(output);
        token.ThrowIfCancellationRequested();
        if (!output.Data.TryGetValue(Mime, out JsonElement resource)) {
            throw new NotSupportedException("The selected output has no saved structured table.");
        }
        if (!TryProperty(resource, "schema", JsonValueKind.Object, out JsonElement schema)
            || !TryProperty(schema, "fields", JsonValueKind.Array, out JsonElement fields)
            || !TryProperty(resource, "data", JsonValueKind.Array, out JsonElement rows)) {
            throw new FormatException("A saved table must contain schema.fields and an inline data array.");
        }

        int columns = fields.GetArrayLength();
        int count = rows.GetArrayLength();
        if (columns == 0) throw new FormatException("The saved table has no columns.");
        if (columns > MaxColumns || count > MaxRows || (long)columns * count > MaxCells) {
            throw new InvalidDataException("The saved table exceeds the export row, column or cell budget.");
        }

        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var positions = new Dictionary<string, int>(columns, StringComparer.Ordinal);
        var headers = new string[columns];
        int index = 0;
        foreach (JsonElement field in fields.EnumerateArray()) {
            if (!TryProperty(field, "name", JsonValueKind.String, out JsonElement name)) {
                throw new FormatException("Every saved table field needs a string name.");
            }
            string header = name.GetString()!;
            if (string.IsNullOrWhiteSpace(header) || !names.Add(header) || header.Length > 255) {
                throw new FormatException("Saved table field names must be nonempty, unique and at most 255 characters.");
            }
            positions.Add(header, index);
            headers[index++] = header;
        }

        var present = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonElement row in rows.EnumerateArray()) {
            token.ThrowIfCancellationRequested();
            if (row.ValueKind != JsonValueKind.Object) throw new FormatException("Saved table rows must be JSON objects.");
            present.Clear();
            foreach (JsonProperty property in row.EnumerateObject()) {
                string name = property.Name;
                if (!present.Add(name) || !positions.ContainsKey(name)) {
                    throw new FormatException("Each saved table row must use unique field names from its schema.");
                }
                if (property.Value.ValueKind is JsonValueKind.Object or JsonValueKind.Array) {
                    throw new NotSupportedException("A spreadsheet cell requires a scalar value.");
                }
            }
        }
        return new NotebookTable(headers, positions, rows);
    }

    internal object?[] ReadRow(JsonElement row, bool excel) {
        var values = new object?[Headers.Length];
        foreach (JsonProperty property in row.EnumerateObject()) {
            values[_columns[property.Name]] = GetValue(property.Value, excel);
        }
        return values;
    }

    private static object? GetValue(JsonElement value, bool excel) {
        switch (value.ValueKind) {
            case JsonValueKind.String: return value.GetString();
            case JsonValueKind.True: return true;
            case JsonValueKind.False: return false;
            case JsonValueKind.Number:
                string number = value.GetRawText();
                if (!excel) return new CsvNumericText(number);
                int digits = SignificantDigits(number);
                if (digits > 15) return number;
                if (value.TryGetInt64(out long integer)) return integer;
                if (digits == 0) return 0L;
                if (value.TryGetDouble(out double floating)
                    && Math.Abs(floating) is >= MinExcelMagnitude and <= MaxExcelMagnitude) {
                    return floating;
                }
                return number;
            default: return null;
        }
    }

    private static int SignificantDigits(string value) {
        int digits = 0;
        int significant = 0;
        foreach (char character in value) {
            if (character is 'e' or 'E') break;
            if (character is < '0' or > '9') continue;
            if (digits > 0 || character != '0') digits++;
            if (character != '0') significant = digits;
        }
        return significant;
    }

    private static bool TryProperty(JsonElement element, string key, JsonValueKind kind, out JsonElement value) {
        value = default;
        return element.ValueKind == JsonValueKind.Object && element.TryGetProperty(key, out value) && value.ValueKind == kind;
    }
}
