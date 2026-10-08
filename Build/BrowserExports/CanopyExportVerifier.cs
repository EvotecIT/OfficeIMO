using System.Data.Common;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;

/// <summary>Independently read every captured cell, order and typed date, and validate XLSX schema/PDF links.</summary>
internal static class CanopyExportVerifier {
    internal static object Verify(string path) {
        using var expected = JsonDocument.Parse(File.ReadAllText(path + ".json"));
        JsonElement contract = expected.RootElement;
        JsonElement[] rows = contract.GetProperty("rows").EnumerateArray().ToArray(), columns = contract.GetProperty("columns").EnumerateArray().ToArray();
        string values = contract.GetProperty("values").GetString()!;
        if (path.EndsWith(".pdf", StringComparison.Ordinal)) {
            var pdf = PdfReadDocument.Open(path); WorkbookVerifier.Require(!pdf.RepairReport.HasRepairs, "Canopy PDF requires repairs.");
            string text = string.Concat(pdf.Pages.SelectMany(page => page.GetTextSpans()).Select(span => span.Text));
            // Remove repeated headings/page labels separately; body values must occur in captured row/column order.
            int position = 0;
            foreach (var row in rows) foreach (var column in columns) {
                string wanted = row.GetProperty("cells").GetProperty(column.GetProperty("id").GetString()!).GetProperty("text").GetString()!;
                foreach (string line in wanted.Replace("\r", "", StringComparison.Ordinal).Split('\n')) {
                    if (line.Length == 0) continue;
                    int found = text.IndexOf(line, position, StringComparison.Ordinal);
                    WorkbookVerifier.Require(found >= position, "Canopy PDF lost or reordered display text: " + line); position = found + line.Length;
                }
            }
            string[] targets = rows.SelectMany(row => columns.Select(column => row.GetProperty("cells").GetProperty(column.GetProperty("id").GetString()!)))
                .Where(cell => cell.TryGetProperty("link", out _)).Select(cell => cell.GetProperty("link").GetProperty("target").GetString()!).ToArray();
            var actual = pdf.Pages.SelectMany(page => page.GetLinkAnnotations()).ToArray();
            WorkbookVerifier.Require(actual.Select(link => link.Uri).SequenceEqual(targets), "Canopy PDF link actions differ.");
            return new { file = Path.GetFileName(path), rows = rows.Length, columns = columns.Length, pages = pdf.Pages.Count, links = actual.Length, repairs = 0 };
        }
        bool excel = path.EndsWith(".xlsx", StringComparison.Ordinal);
        if (excel) WorkbookVerifier.Verify(path);
        using DbDataReader reader = excel ? ExcelDocument.OpenDataReader(path, new ExcelReadOptions { HasHeaderRow = true }) : CsvDocument.OpenDataReader(path);
        WorkbookVerifier.Require(reader.FieldCount == columns.Length, "Canopy exported column count differs.");
        for (int c = 0; c < columns.Length; c++) WorkbookVerifier.Require(reader.GetName(c) == columns[c].GetProperty("title").GetString(), "Canopy exported column order differs.");
        foreach (JsonElement row in rows) {
            WorkbookVerifier.Require(reader.Read(), "Canopy export lost a record.");
            for (int c = 0; c < columns.Length; c++) {
                JsonElement cell = row.GetProperty("cells").GetProperty(columns[c].GetProperty("id").GetString()!);
                JsonElement wanted = cell.GetProperty(values == "display" ? "text" : "value");
                object actual = reader.GetValue(c);
                bool equal;
                if (!excel) {
                    string text = wanted.ValueKind == JsonValueKind.Null ? "" : wanted.ValueKind == JsonValueKind.String ? wanted.GetString()! : wanted.ValueKind == JsonValueKind.True ? "True" : wanted.ValueKind == JsonValueKind.False ? "False" : wanted.GetRawText();
                    if (wanted.ValueKind == JsonValueKind.String && Regex.IsMatch(text, @"^[\s\uFEFF]*[=+\-@]")) text = "'" + text;
                    equal = text == Convert.ToString(actual, CultureInfo.InvariantCulture);
                } else if (wanted.ValueKind == JsonValueKind.String && values == "raw" && columns[c].GetProperty("kind").GetString() == "datetime" &&
                    !Regex.IsMatch(wanted.GetString()!, @"\.\d{3}\d*[1-9]\d*(?:Z|[+-])") && DateTimeOffset.Parse(wanted.GetString()!, CultureInfo.InvariantCulture).Year >= 1900) {
                    DateTime date = DateTimeOffset.Parse(wanted.GetString()!, CultureInfo.InvariantCulture).UtcDateTime;
                    equal = actual is DateTime timestamp && Math.Abs((timestamp - DateTime.SpecifyKind(date, DateTimeKind.Unspecified)).TotalMilliseconds) <= 1;
                } else equal = wanted.ValueKind switch {
                    JsonValueKind.Null => actual == DBNull.Value || actual is null || actual is string { Length: 0 },
                    JsonValueKind.Number => actual is double or int or long or decimal && Math.Abs(Convert.ToDouble(actual, CultureInfo.InvariantCulture) - wanted.GetDouble()) < 1e-9,
                    JsonValueKind.True => Equals(actual, true), JsonValueKind.False => Equals(actual, false),
                    _ => Equals(actual, wanted.GetString())
                };
                WorkbookVerifier.Require(equal, $"Canopy cell {row.GetProperty("id").GetString()}/{c} differs: {actual}.");
            }
        }
        WorkbookVerifier.Require(!reader.Read(), "Canopy export added a record.");
        if (excel) {
            using var sdk = SpreadsheetDocument.Open(path, false);
            var part = sdk.WorkbookPart!.WorksheetParts.Single();
            string[] targets = rows.SelectMany(row => columns.Select(column => row.GetProperty("cells").GetProperty(column.GetProperty("id").GetString()!)))
                .Where(cell => cell.TryGetProperty("link", out _)).Select(cell => cell.GetProperty("link").GetProperty("target").GetString()!).ToArray();
            var links = part.Worksheet!.Descendants<Hyperlink>().ToArray();
            WorkbookVerifier.Require(links.Select(link => part.HyperlinkRelationships.Single(r => r.Id == link.Id!.Value).Uri.AbsoluteUri).SequenceEqual(targets), "Canopy XLSX external links differ.");
            var styles = sdk.WorkbookPart.WorkbookStylesPart!.Stylesheet!;
            var body = part.Worksheet.Descendants<Row>().Skip(1).ToArray();
            for (int r = 0; r < rows.Length; r++) for (int c = 0; c < columns.Length; c++) {
                JsonElement cell = rows[r].GetProperty("cells").GetProperty(columns[c].GetProperty("id").GetString()!);
                string? rowTone = rows[r].TryGetProperty("tone", out var rt) ? rt.GetString() : null;
                string? cellTone = cell.TryGetProperty("tone", out var ct) ? ct.GetString() : null;
                string? tone = cellTone is not null and not "neutral" ? cellTone : rowTone;
                string? expectedFill = tone switch { "warning" => "FFFFF2CC", "info" => "FFDDEBF7", "success" => "FFE2F0D9", _ => null };
                var saved = body[r].Elements<Cell>().ElementAt(c);
                var style = styles.CellFormats!.Elements<CellFormat>().ElementAt((int)(saved.StyleIndex?.Value ?? 0));
                var fill = styles.Fills!.Elements<Fill>().ElementAt((int)(style.FillId?.Value ?? 0));
                WorkbookVerifier.Require(fill.PatternFill?.ForegroundColor?.Rgb?.Value == expectedFill, "Canopy cell/row tone precedence differs.");
                string? id = columns[c].GetProperty("id").GetString();
                string? code = styles.NumberingFormats?.Elements<NumberingFormat>().FirstOrDefault(f => f.NumberFormatId?.Value == style.NumberFormatId?.Value)?.FormatCode?.Value;
                if (id == "amount") WorkbookVerifier.Require(code == "0.00", "Canopy number format was lost under presentation.");
                if (id == "when") WorkbookVerifier.Require(code == "yyyy-mm-dd hh:mm:ss.000", "Canopy date format was lost under presentation.");
            }
        }
        return new { file = Path.GetFileName(path), rows = rows.Length, columns = columns.Length, openXml = excel };
    }
}
