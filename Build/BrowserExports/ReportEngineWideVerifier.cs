using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Pdf;

/// <summary>Independent reads of the workbench's sixteen typed fields and its repeated-ID PDF download.</summary>
internal static class ReportEngineWideVerifier {
    private static readonly XNamespace Sheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    internal static object Verify(string path) {
        using var expected = JsonDocument.Parse(File.ReadAllText(path + ".json")); var contract = expected.RootElement;
        return path.EndsWith(".xlsx", StringComparison.Ordinal) ? Workbook(path, contract) : Pdf(path, contract);
    }
    private static object Workbook(string path, JsonElement contract) {
        WorkbookVerifier.Verify(path);
        using var package = SpreadsheetDocument.Open(path, false); var book = package.WorkbookPart!;
        var sheets = book.Workbook!.Sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().ToArray();
        Require(sheets.Length == 3, "Wide workbook lost its context, detail or overview sheet.");
        var detail = (WorksheetPart)book.GetPartById(sheets[1].Id!.Value!);
        using var stream = detail.GetStream(); var xml = XDocument.Load(stream);
        string[] shared = book.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().Select(item => item.InnerText).ToArray() ?? [];
        string Text(XElement? cell) => cell is null ? "" : (string?)cell.Attribute("t") == "s" ? shared[int.Parse(cell.Element(Sheet + "v")!.Value, CultureInfo.InvariantCulture)]
            : (string?)cell.Attribute("t") == "inlineStr" ? string.Concat(cell.Descendants(Sheet + "t").Select(text => text.Value)) : cell.Element(Sheet + "v")?.Value ?? "";
        var columns = contract.GetProperty("columns").EnumerateArray().ToArray(); var rows = contract.GetProperty("rows").EnumerateArray().ToArray();
        var expectedById = rows.ToDictionary(row => row[0].GetString()!, StringComparer.Ordinal); var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var row in xml.Descendants(Sheet + "row")) {
            var cells = row.Elements(Sheet + "c").ToDictionary(cell => Regex.Replace((string)cell.Attribute("r")!, @"\d+$", ""));
            if (!cells.TryGetValue("B", out var service) || !expectedById.TryGetValue(Text(service), out var record)) continue;
            Require(seen.Add(Text(service)), "A selected wide record was duplicated.");
            for (int column = 0; column < columns.Length; column++) {
                cells.TryGetValue(((char)('B' + column)).ToString(), out var cell); string actual = Text(cell); var value = record[column];
                Require(cell?.Element(Sheet + "f") is null, "A selected typed value became a formula.");
                if (value.ValueKind == JsonValueKind.Null) Require(actual.Length == 0, "A null wide field became a value.");
                else if (value.ValueKind == JsonValueKind.Number || columns[column].TryGetProperty("type", out var type) && type.GetString() == "date") {
                    double number = value.ValueKind == JsonValueKind.Number ? value.GetDouble() : DateTimeOffset.Parse(value.GetString()!, CultureInfo.InvariantCulture).UtcDateTime.ToOADate();
                    Require(((string?)cell?.Attribute("t") is null or "n") && double.TryParse(actual, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) && Math.Abs(parsed - number) < 1e-8, "A typed wide number/date differs.");
                } else Require(actual == value.GetString(), "A selected wide text field differs.");
            }
        }
        Require(seen.Count == rows.Length && columns.Length == 16, "Wide workbook record or visible-column count differs.");
        return new { file = Path.GetFileName(path), format = "xlsx", rows = seen.Count, typedColumns = columns.Length, openXml = true, bytes = new FileInfo(path).Length };
    }
    private static object Pdf(string path, JsonElement contract) {
        using var document = PdfDocument.Load(path, new PdfLoadOptions { IncludeArtifactText = true });
        Require(!document.RepairReport.HasRepairs, "Wide PDF required repair.");
        string all = string.Concat(document.Pages.SelectMany(page => page.GetTextSpans()).Select(span => span.Text)), compact = Compact(all);
        foreach (var column in contract.GetProperty("columns").EnumerateArray()) Require(compact.Contains(Compact(column.GetProperty("header").GetString()!), StringComparison.Ordinal), "Wide PDF lost a selected heading.");
        foreach (string key in new[] { "scope", "title", "ending", "advice" }) Require(compact.Contains(Compact(contract.GetProperty(key).GetString()!), StringComparison.Ordinal), "Wide PDF lost composed report context.");
        foreach (string value in contract.GetProperty("display").EnumerateArray().SelectMany(row => row.EnumerateArray()).Select(value => value.GetString()!).Where(value => value.Length > 0).Distinct(StringComparer.Ordinal))
            Require(compact.Contains(Compact(value), StringComparison.Ordinal), "Wide PDF lost a localized selected value: " + value);
        var ids = contract.GetProperty("rows").EnumerateArray().Select(row => row[0].GetString()!).ToArray();
        var actual = Regex.Matches(all, @"WEB-\d{6}").Select(match => match.Value).ToArray();
        int bands = ids.Length == 0 ? 0 : actual.Count(id => id == ids[0]);
        Require(ids.Length == 0 ? actual.Length == 0 : bands >= 2 && ids.All(id => actual.Count(value => value == id) == bands) && actual.All(id => ids.Contains(id, StringComparer.Ordinal)), "Wide PDF identifiers were invented, lost or repeated unevenly across bands.");
        return new { file = Path.GetFileName(path), format = "pdf", rows = ids.Length, typedColumns = 16, bands, pages = document.Pages.Count, repairs = 0,
            boundary = "Actual workbench download; shared host fixtures separately verify every band value, window and point width.", bytes = new FileInfo(path).Length };
    }
    private static string Compact(string value) => Regex.Replace(value, @"\s+", "");
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidDataException("Wide workbench artifact: " + message); }
}
