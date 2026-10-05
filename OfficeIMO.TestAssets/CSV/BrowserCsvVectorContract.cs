#if NET8_0_OR_GREATER
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using OfficeIMO.CSV;

namespace OfficeIMO.TestAssets;

// Shared by the ordinary CSV tests and the opt-in browser interoperability runner.
internal static class BrowserCsvVectorContract {
    internal static byte[] Write(JsonElement vector) {
        var document = new CsvDocument().WithHeader(vector.GetProperty("columns").EnumerateArray()
            .Select(column => column.GetProperty("header").GetString()!).ToArray());
        foreach (JsonElement row in vector.GetProperty("rows").EnumerateArray())
            document.AddRow(row.EnumerateArray().Select(Value).ToArray());
        var options = new CsvSaveOptions {
            Delimiter = vector.TryGetProperty("delimiter", out JsonElement delimiter) ? delimiter.GetString()![0] : ',',
            NewLine = vector.TryGetProperty("lineEnding", out JsonElement lineEnding) ? lineEnding.GetString()! : "\r\n",
            IncludeHeader = !vector.TryGetProperty("includeHeader", out JsonElement header) || header.GetBoolean(),
            Encoding = new UTF8Encoding(vector.TryGetProperty("bom", out JsonElement bom) && bom.GetBoolean()),
            Culture = CultureInfo.InvariantCulture,
            DateTimeFormat = "yyyy-MM-dd'T'HH:mm:ss.fff'Z'",
            UseUtc = true,
            FormulaInjectionPolicy = vector.TryGetProperty("formulaInjectionProtection", out JsonElement protect) && !protect.GetBoolean()
                ? CsvFormulaInjectionPolicy.Preserve : CsvFormulaInjectionPolicy.Escape
        };
        return document.ToBytes(options);
    }

    internal static byte[] Expected(JsonElement vector) {
        var encoding = new UTF8Encoding(vector.TryGetProperty("bom", out JsonElement bom) && bom.GetBoolean());
        return encoding.GetPreamble().Concat(encoding.GetBytes(vector.GetProperty("expected").GetString()!)).ToArray();
    }

    private static object? Value(JsonElement value) => value.ValueKind switch {
        JsonValueKind.Null => null,
        JsonValueKind.String => value.GetString(),
        JsonValueKind.Number => value.GetDouble(),
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.Object when value.GetProperty("kind").GetString() == "date" =>
            DateTime.Parse(value.GetProperty("value").GetString()!, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal),
        _ => throw new InvalidDataException("Unsupported shared CSV vector value.")
    };
}
#endif
