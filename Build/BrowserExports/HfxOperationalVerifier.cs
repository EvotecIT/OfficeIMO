using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Pdf;

/// <summary>Checks actual generated DOM capture and exported files against C# source-model expectations.</summary>
internal static class HfxOperationalVerifier {
    private static readonly XNamespace Sheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>Compares captured DOM values and hidden group paths with the C# model, before exporting.</summary>
    internal static void VerifyCapture(JsonElement model, JsonElement expected, JsonElement actual) {
        string[] keys = Strings(expected.GetProperty("keys"));
        Require(Strings(actual.GetProperty("keys")).SequenceEqual(keys), "Visible stable field order differs.");
        Require(Strings(actual.GetProperty("headers")).SequenceEqual(Strings(expected.GetProperty("headers"))), "Captured display headings differ.");
        JsonElement[] rows = expected.GetProperty("rows").EnumerateArray().ToArray(), captured = actual.GetProperty("rows").EnumerateArray().ToArray();
        Require(Strings(actual.GetProperty("ids")).SequenceEqual(rows.Select(row => row.GetProperty("id").GetString()!)), "Selected/filter/reordered row identities differ.");
        Require(captured.Length == rows.Length, "Captured row count differs.");
        for (int index = 0; index < rows.Length; index++) {
            JsonElement[] values = rows[index].GetProperty("values").EnumerateArray().ToArray();
            Require(captured[index].GetArrayLength() == keys.Length + 3, "Capture lost hidden grouping metadata.");
            for (int column = 0; column < keys.Length; column++) {
                JsonElement value = captured[index][column], wanted = values[column];
                string type = Type(model, keys[column]);
                Require(wanted.ValueKind == JsonValueKind.Null ? value.ValueKind == JsonValueKind.Null
                    : type == "date" ? DateTimeOffset.Parse(value.GetString()!, CultureInfo.InvariantCulture) == DateTimeOffset.Parse(wanted.GetString()!, CultureInfo.InvariantCulture)
                    : wanted.ValueKind == JsonValueKind.Number ? value.ValueKind == JsonValueKind.Number && value.GetDouble() == wanted.GetDouble()
                    : value.GetString() == wanted.GetString(), "Captured raw value/type differs at " + index + "/" + keys[column] + ".");
            }
            Require(captured[index].EnumerateArray().Skip(keys.Length).Select(value => value.GetString()!).SequenceEqual(Strings(rows[index].GetProperty("path"))), "Captured group path differs.");
        }
    }

    /// <summary>Checks selected identity, accepted bytes and bounded worker-message/listener ownership.</summary>
    internal static void VerifyCompletion(JsonElement expected, JsonElement actual, long bytes) {
        Require(!actual.TryGetProperty("error", out _), "Worker export failed.");
        Require(Strings(actual.GetProperty("ids")).SequenceEqual(expected.GetProperty("rows").EnumerateArray().Select(row => row.GetProperty("id").GetString()!)), "Worker source scope differs from capture.");
        Require(Strings(actual.GetProperty("keys")).SequenceEqual(Strings(expected.GetProperty("keys"))), "Worker selected column order differs.");
        Require(actual.GetProperty("completed").GetProperty("rows").GetInt32() == expected.GetProperty("rows").GetArrayLength()
            && actual.GetProperty("completed").GetProperty("bytes").GetInt64() == bytes, "Worker completion rows/bytes differ.");
        Require(actual.GetProperty("observed").GetProperty("maxRows").GetInt32() <= 3
            && actual.GetProperty("observed").GetProperty("maxChunkBytes").GetInt32() <= 4096 && actual.GetProperty("remainingListeners").GetInt32() == 0,
            "Worker transport exceeded bounds or retained endpoint listeners.");
    }

    /// <summary>Reads the delivered workbook or PDF independently; does not claim rendered/native acceptance.</summary>
    internal static object Verify(string path, JsonElement model, JsonElement expected) => path.EndsWith(".xlsx", StringComparison.Ordinal)
        ? VerifyWorkbook(path, model, expected) : VerifyPdf(path, expected);

    private static object VerifyWorkbook(string path, JsonElement model, JsonElement expected) {
        WorkbookVerifier.Verify(path);
        using var package = SpreadsheetDocument.Open(path, false);
        WorkbookPart book = package.WorkbookPart!;
        var sheets = book.Workbook!.Sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().ToArray();
        Require(sheets.Select(sheet => sheet.Name!.Value).SequenceEqual(new[] { "Report context", "Selected inventory" }), "Workbook context/analysis sheet order differs.");
        var sheet = sheets.Single(sheet => sheet.Name!.Value == "Selected inventory");
        var part = (WorksheetPart)book.GetPartById(sheet.Id!.Value!);
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read); var xml = XDocument.Load(stream);
        string[] shared = book.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().Select(item => item.InnerText).ToArray() ?? [];
        string Text(XElement? cell) => cell is null ? "" : (string?)cell.Attribute("t") == "s" ? shared[int.Parse(cell.Element(Sheet + "v")!.Value, CultureInfo.InvariantCulture)]
            : (string?)cell.Attribute("t") == "inlineStr" ? string.Concat(cell.Descendants(Sheet + "t").Select(value => value.Value)) : cell.Element(Sheet + "v")?.Value ?? "";
        var contextPart = (WorksheetPart)book.GetPartById(sheets[0].Id!.Value!);
        using var contextStream = contextPart.GetStream(FileMode.Open, FileAccess.Read);
        var contextXml = XDocument.Load(contextStream);
        string contextText = string.Join("\n", contextXml.Descendants(Sheet + "c").Select(Text));
        Require(contextText.Contains("Operational inventory review", StringComparison.Ordinal)
            && contextText.Contains("Counts in group labels describe the exported selection", StringComparison.Ordinal), "Workbook lacks its report scope/count interpretation.");
        string[] keys = Strings(expected.GetProperty("keys")), headers = Strings(expected.GetProperty("headers"));
        foreach (string key in new[] { "latency", "availability", "requests" }.Intersect(keys, StringComparer.Ordinal)) {
            string statistic = key == "requests" ? "sum" : "average";
            Require(contextText.Contains(headers[Array.IndexOf(keys, key)] + " — " + statistic, StringComparison.Ordinal), "Workbook summary operation legend differs from the selected metric.");
        }
        JsonElement[] wanted = expected.GetProperty("rows").EnumerateArray().ToArray();
        int serviceColumn = Array.IndexOf(keys, "service") + 2;
        XElement? Cell(XElement row, int column) => row.Elements(Sheet + "c").FirstOrDefault(cell => (string?)cell.Attribute("r") == Column(column) + row.Attribute("r")!.Value);
        XElement[] rows = xml.Descendants(Sheet + "row").ToArray();
        Require(rows.All(row => row.Elements(Sheet + "c").Count() <= keys.Length + 1), "Workbook gained unselected or grouping-metadata columns.");
        XElement heading = rows.Single(row => (string?)row.Attribute("r") == "2");
        Require(headers.Select((header, index) => Text(Cell(heading, index + 2)) == header).All(value => value), "Workbook visible heading order differs.");
        XElement[] detail = rows.Where(row => Regex.IsMatch(Text(Cell(row, serviceColumn)), "^OPS-[0-9]{4}$", RegexOptions.CultureInvariant)).ToArray();
        Require(detail.Length == wanted.Length, "Workbook selected row count differs.");
        for (int row = 0; row < wanted.Length; row++) for (int column = 0; column < keys.Length; column++) {
            XElement? cell = Cell(detail[row], column + 2); JsonElement value = wanted[row].GetProperty("values")[column];
            Require(cell?.Element(Sheet + "f") is null, "An inventory literal became a formula.");
            if (value.ValueKind == JsonValueKind.Null) Require(Text(cell) == "", "A selected null became zero or text.");
            else if (value.ValueKind == JsonValueKind.Number || Type(model, keys[column]) == "date") {
                double number = value.ValueKind == JsonValueKind.Number ? value.GetDouble()
                    : (DateTimeOffset.Parse(value.GetString()!, CultureInfo.InvariantCulture).UtcDateTime - new DateTime(1899, 12, 30)).TotalDays;
                Require((string?)cell?.Attribute("t") is null or "n" && double.TryParse(Text(cell), NumberStyles.Float, CultureInfo.InvariantCulture, out double actual)
                    && Math.Abs(actual - number) < 1e-9, "A selected number/date type or value differs.");
            } else Require(Text(cell) == value.GetString(), "Selected text or row order differs.");
        }
        XElement total = rows.Single(row => Text(Cell(row, 1)) == "All selected services · " + wanted.Length);
        foreach (string key in new[] { "latency", "availability", "requests" }.Intersect(keys, StringComparer.Ordinal)) {
            int index = Array.IndexOf(keys, key);
            double[] values = wanted.Select(row => row.GetProperty("values")[index]).Where(value => value.ValueKind == JsonValueKind.Number).Select(value => value.GetDouble()).ToArray();
            XElement? cell = Cell(total, index + 2);
            if (values.Length == 0) Require(Text(cell) == (key == "requests" ? "0" : ""), "Empty sum must be zero and empty averages must remain blank.");
            else Require(cell?.Element(Sheet + "f") is not null && double.TryParse(Text(cell), NumberStyles.Float, CultureInfo.InvariantCulture, out double actual)
                && Math.Abs(actual - (key == "requests" ? values.Sum() : values.Average())) < 1e-9, "Native total formula/cache differs from selected C# values.");
        }
        return new { format = "xlsx", rows = wanted.Length, columns = keys.Length, independentSelectedValues = true, openXml = true, readers = 2 };
    }

    private static object VerifyPdf(string path, JsonElement expected) {
        using var document = PdfDocument.Load(path);
        Require(!document.RepairReport.HasRepairs, "PDF required reader repairs.");
        string text = string.Concat(document.Pages.SelectMany(page => page.GetTextSpans()).Select(span => span.Text));
        JsonElement[] rows = expected.GetProperty("rows").EnumerateArray().ToArray();
        Require(Regex.Matches(text, "OPS-[0-9]{4}", RegexOptions.CultureInvariant).Select(match => match.Value)
            .SequenceEqual(rows.Select(row => row.GetProperty("id").GetString()!)), "Logical PDF service IDs were lost, duplicated, omitted from selection or reordered.");
        string compact = Regex.Replace(text, @"\s+", "");
        foreach (string header in Strings(expected.GetProperty("headers"))) Require(compact.Contains(Regex.Replace(header, @"\s+", ""), StringComparison.Ordinal), "PDF lost a selected visible heading.");
        Require(compact.Contains("Allselectedservices·" + rows.Length, StringComparison.Ordinal), "PDF total count differs.");
        if (rows.Any(row => row.GetProperty("id").GetString() == "OPS-0085")) Require(text.Contains("OPSNOTE084-035", StringComparison.Ordinal), "PDF lost the final long-value fragment.");
        return new { format = "pdf", rows = rows.Length, pages = document.Pages.Count, logicalSelectedOrder = true, repairs = 0,
            limits = "Structural/text readback; rendered/native/AT acceptance remains separate" };
    }

    private static string Type(JsonElement model, string key) => model.GetProperty("columns").EnumerateArray().Single(column => column.GetProperty("key").GetString() == key).GetProperty("type").GetString()!;
    private static string[] Strings(JsonElement value) => value.EnumerateArray().Select(item => item.GetString()!).ToArray();
    private static string Column(int ordinal) { string result = ""; while (ordinal > 0) { ordinal--; result = (char)('A' + ordinal % 26) + result; ordinal /= 26; } return result; }
    private static void Require(bool value, string message) => WorkbookVerifier.Require(value, "Actual HFX inventory: " + message);
}
