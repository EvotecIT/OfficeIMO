using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

/// <summary>Independent, validation-only entrypoint for opt-in cross-library table artifacts.</summary>
internal static class TableEngineComparisonVerifier {
    internal static object Verify(string[] args) {
        if (args.Length < 2 || !Path.IsPathRooted(args[1]))
            throw new ArgumentException("Supply an absolute table artifact path and explicit fixture options.");
        string path = Path.GetFullPath(args[1]);
        var options = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (string argument in args.Skip(2)) {
            int equals = argument.IndexOf('=');
            if (!argument.StartsWith("--", StringComparison.Ordinal) || equals < 3
                || !options.TryAdd(argument[2..equals], argument[(equals + 1)..]))
                throw new ArgumentException("Invalid or duplicate comparison option: " + argument);
        }
        string[] allowed = ["format", "rows", "columns", "unique", "text-profile"];
        if (options.Keys.Any(key => !allowed.Contains(key, StringComparer.Ordinal)))
            throw new ArgumentException("Unknown table comparison option.");
        string format = Required("format"), textProfile = Required("text-profile");
        int rows = Positive("rows"), columns = Positive("columns");
        if (format is not ("csv" or "xlsx" or "pdf") || textProfile is not ("unicode" or "bmp") || columns > 16384
            || !bool.TryParse(Required("unique"), out bool unique))
            throw new ArgumentException("Unsupported table comparison fixture.");
        object proof = format == "pdf" ? DataTablesPdfComparisonVerifier.Verify(path, rows, columns, unique)
            : DataTablesComparisonVerifier.Verify(path, format, rows, columns, unique, textProfile: textProfile);
        object conformance;
        if (format == "xlsx") {
            using SpreadsheetDocument document = SpreadsheetDocument.Open(path, false);
            string[] errors = new OpenXmlValidator().Validate(document).Take(20).Select(error => error.Description).ToArray();
            if (errors.Length != 0) throw new InvalidDataException("XLSX schema validation failed: " + string.Join("; ", errors));
            conformance = new { passed = true, errors };
        } else conformance = format == "csv"
            ? new { passed = true, quoteAll = true, delimiter = ",", lineEnding = "CRLF", bom = false }
            : (object)new { passed = true, notApplicable = true };
        using var input = File.OpenRead(path);
        string sha256 = Convert.ToHexString(SHA256.HashData(input)).ToLowerInvariant();
        if (format == "csv" && sha256 != QuotedCsvHash(rows, columns, unique, textProfile))
            throw new InvalidDataException("CSV bytes must use quoted fields, commas, CRLF and no BOM.");
        return new { passed = true, sha256,
            format, rows, columns, proof, conformance };

        string Required(string key) => options.TryGetValue(key, out string? value) && value.Length != 0 ? value
            : throw new ArgumentException("Missing table comparison option: --" + key);
        int Positive(string key) => int.TryParse(Required(key), NumberStyles.None, CultureInfo.InvariantCulture, out int value) && value > 0
            ? value : throw new ArgumentException("Table comparison dimensions must be positive integers.");
    }

    private static string QuotedCsvHash(int rows, int columns, bool unique, string textProfile) {
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        for (int row = -1; row < rows; row++) {
            var line = new StringBuilder();
            for (int column = 0; column < columns; column++) {
                if (column != 0) line.Append(',');
                string value = row < 0 ? "Column " + (column + 1)
                    : DataTablesComparisonVerifier.Expected(row, column, unique, textProfile);
                line.Append('"').Append(value.Replace("\"", "\"\"")).Append('"');
            }
            line.Append("\r\n");
            hash.AppendData(Encoding.UTF8.GetBytes(line.ToString()));
        }
        return Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant();
    }
}
