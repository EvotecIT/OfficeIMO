using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Pdf;

/// <summary>Read grouped captures independently, including numeric/null totals and native finite formulas.</summary>
internal static class GroupedReportVerifier {
    private static readonly XNamespace Sheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    internal static object Verify(string path) {
        using var document = JsonDocument.Parse(File.ReadAllText(path + ".json"));
        return path.EndsWith(".xlsx", StringComparison.Ordinal) ? VerifyWorkbook(path, document.RootElement) : VerifyPdf(path, document.RootElement);
    }

    private static object VerifyWorkbook(string path, JsonElement contract) {
        WorkbookVerifier.Verify(path);
        using var package = SpreadsheetDocument.Open(path, false);
        WorkbookPart book = package.WorkbookPart!;
        var sheet = book.Workbook!.Sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Single();
        Require(sheet.State is null || sheet.State.Value == SheetStateValues.Visible, "The grouped report has hidden data.");
        var part = (WorksheetPart)book.GetPartById(sheet.Id!.Value!);
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        var xml = XDocument.Load(stream);
        Require(!part.TableDefinitionParts.Any() && !xml.Descendants(Sheet + "autoFilter").Any(), "Grouped summaries were placed inside a table/autofilter.");
        string[] shared = book.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().Select(item => item.InnerText).ToArray() ?? [];
        var cells = xml.Descendants(Sheet + "c").ToDictionary(cell => cell.Attribute("r")!.Value, StringComparer.Ordinal);
        string Text(XElement cell) => (string?)cell.Attribute("t") == "s" ? shared[int.Parse(cell.Element(Sheet + "v")!.Value, CultureInfo.InvariantCulture)]
            : (string?)cell.Attribute("t") == "inlineStr" ? string.Concat(cell.Descendants(Sheet + "t").Select(text => text.Value)) : cell.Element(Sheet + "v")?.Value ?? "";
        XElement Cell(int row, int column) => cells.TryGetValue(Column(column) + row.ToString(CultureInfo.InvariantCulture), out var cell)
            ? cell : throw new InvalidDataException("Grouped cell missing at " + Column(column) + row + ".");
        void Value(int row, int column, JsonElement value, string key) {
            XElement cell = Cell(row, column); Require(cell.Element(Sheet + "f") is null, "A captured value became a formula.");
            if (value.ValueKind == JsonValueKind.Null) Require(Text(cell) == "", "Null metrics became text or zero.");
            else if (value.ValueKind == JsonValueKind.Number) Number(cell, value.GetDouble());
            else if (key == "checked") Number(cell, (DateTimeOffset.Parse(value.GetString()!, CultureInfo.InvariantCulture).UtcDateTime - new DateTime(1899, 12, 30)).TotalDays);
            else Require(Text(cell) == value.GetString(), "Raw captured text/order differs at " + Column(column) + row + ".");
        }
        string[] keys = contract.GetProperty("keys").EnumerateArray().Select(value => value.GetString()!).ToArray();
        string[] headers = contract.GetProperty("headers").EnumerateArray().Select(value => value.GetString()!).ToArray();
        JsonElement[] rows = contract.GetProperty("rows").EnumerateArray().ToArray(), groups = contract.GetProperty("groups").EnumerateArray().ToArray();
        Require(Text(Cell(2, 1)) == contract.GetProperty("labelHeader").GetString(), "Leading host-authored heading differs.");
        for (int column = 0; column < headers.Length; column++) Require(Text(Cell(2, column + 2)) == headers[column], "Translated/custom heading or current column order differs.");
        int output = 3, source = 0, firstDetail = 0, lastDetail = 0, formulaCount = 0;
        foreach (JsonElement group in groups) {
            Require(Text(Cell(output++, 1)) == group.GetProperty("label").GetString(), "Host-authored group label differs.");
            int first = output, count = group.GetProperty("rows").GetInt32(); firstDetail = firstDetail == 0 ? first : firstDetail;
            for (int index = 0; index < count; index++, output++, source++) {
                var row = xml.Descendants(Sheet + "row").Single(item => (int)item.Attribute("r")! == output);
                Require((string?)row.Attribute("outlineLevel") == "1" && (string?)row.Attribute("hidden") is not "1" and not "true", "Detail rows lost their editable expanded outline.");
                for (int column = 0; column < keys.Length; column++) Value(output, column + 2, rows[source][column], keys[column]);
            }
            lastDetail = output - 1;
            Require(Text(Cell(output, 1)) == group.GetProperty("subtotal").GetString(), "Group count/label differs.");
            for (int column = 0; column < keys.Length; column++) if (keys[column] is "latency" or "availability") {
                Formula(Cell(output, column + 2), Column(column + 2), first, lastDetail, group.GetProperty(keys[column])); formulaCount++;
            }
            output++;
        }
        Require(source == rows.Length && Text(Cell(output, 1)) == contract.GetProperty("totalLabel").GetString(), "Grand total count or selected row count differs.");
        for (int column = 0; column < keys.Length; column++) if (keys[column] is "latency" or "availability") {
            JsonElement total = contract.GetProperty("totals").GetProperty(keys[column]);
            if (rows.Length == 0) Require(Text(Cell(output, column + 2)) == "" && Cell(output, column + 2).Element(Sheet + "f") is null, "Empty average must remain blank.");
            else { Formula(Cell(output, column + 2), Column(column + 2), firstDetail, lastDetail, total); formulaCount++; }
        }
        Require(xml.Descendants(Sheet + "row").Count() == output, "Unexpected body/summary rows were added or lost.");
        Require(xml.Descendants(Sheet + "f").Count() == formulaCount, "Unexpected formulas outside authored group summaries.");
        Require(!string.Concat(cells.Values.Select(Text)).Contains("INTERNAL-ONLY", StringComparison.Ordinal) && !string.Concat(cells.Values.Select(Text)).Contains("UNSELECTED-", StringComparison.Ordinal), "Excluded source data leaked.");
        return new { file = Path.GetFileName(path), format = "xlsx", rows = rows.Length, groups = groups.Length, formulas = formulaCount, visibleSheets = 1,
            openXml = true, readers = 2, bytes = new FileInfo(path).Length };
    }

    private static void Formula(XElement cell, string letter, int first, int last, JsonElement expected) {
        string range = letter + first + ":" + letter + last;
        Require(cell.Element(Sheet + "f")?.Value == "IF(SUBTOTAL(2," + range + ")=0,\"\",SUBTOTAL(1," + range + "))", "Summary lost its finite native SUBTOTAL average range.");
        if (expected.ValueKind == JsonValueKind.Null) Require(cell.Element(Sheet + "v")?.Value is null or "", "Empty average gained a numeric formula cache.");
        else Number(cell, expected.GetDouble());
    }

    private static void Number(XElement cell, double expected) {
        Require((string?)cell.Attribute("t") is null or "n", "A raw numeric/date value or formula cache became text.");
        Require(double.TryParse(cell.Element(Sheet + "v")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double actual) && double.IsFinite(actual) && Math.Abs(actual - expected) < 1e-9,
            "Raw number/date or independent average cache differs.");
    }

    private static object VerifyPdf(string path, JsonElement contract) {
        var document = PdfReadDocument.Open(path);
        Require(!document.RepairReport.HasRepairs, "Grouped PDF required repairs.");
        string[] pageTexts = document.Pages.Select(page => string.Concat(page.GetTextSpans().Select(span => span.Text))).ToArray();
        string text = string.Concat(pageTexts);
        JsonElement[] rows = contract.GetProperty("rows").EnumerateArray().ToArray(), groups = contract.GetProperty("groups").EnumerateArray().ToArray();
        string[] keys = contract.GetProperty("keys").EnumerateArray().Select(value => value.GetString()!).ToArray();
        int service = Array.IndexOf(keys, "service");
        Require(Regex.Matches(text, @"WEB-\d{6}").Select(match => match.Value).SequenceEqual(rows.Select(row => row[service].GetString()!)), "PDF selected IDs were lost, duplicated or reordered.");
        Require(!text.Contains("INTERNAL-ONLY", StringComparison.Ordinal) && !text.Contains("UNSELECTED-", StringComparison.Ordinal) && !text.Contains("DISPLAY-ONLY", StringComparison.Ordinal), "Excluded/display-only grid values leaked into PDF.");
        foreach (JsonElement group in groups) {
            Require(text.Contains(group.GetProperty("label").GetString()!, StringComparison.Ordinal), "PDF group label differs.");
            string label = group.GetProperty("subtotal").GetString()!;
            Require(Compact(text).Contains(Compact(label), StringComparison.Ordinal), "PDF group subtotal count/label differs.");
        }
        bool includeGrandTotal = !contract.TryGetProperty("includeGrandTotal", out var grandFlag) || grandFlag.GetBoolean();
        bool hasGrandLabel = Compact(text).Contains(Compact(contract.GetProperty("totalLabel").GetString()!), StringComparison.Ordinal);
        Require(includeGrandTotal ? hasGrandLabel : !hasGrandLabel, "PDF grand total label/count differs from the requested summary scope.");
        string compactText = Compact(text);
        foreach (JsonElement summary in contract.GetProperty("summaryRows").EnumerateArray()) {
            string label = Compact(summary.GetProperty("label").GetString()!);
            string values = Compact(string.Concat(summary.GetProperty("values").EnumerateArray().Select(value => value.GetString())));
            Require(compactText.Contains(label + values, StringComparison.Ordinal), "PDF independent summary averages differ from the selected raw source.");
        }
        Require(groups.Length == 0 || document.Pages.Count > groups.Length, "The fixture did not exercise group continuation across pages.");
        string suffix = contract.GetProperty("continuationSuffix").GetString()!;
        Require(groups.Length == 0 || text.Contains(suffix, StringComparison.Ordinal), "Continued group pages lost their explicit host label.");
        string[] headers = contract.GetProperty("headers").EnumerateArray().Select(value => value.GetString()!).ToArray();
        foreach (string heading in headers.Prepend(contract.GetProperty("labelHeader").GetString()!))
            Require(Compact(text).Contains(Compact(heading), StringComparison.Ordinal), "PDF translated/custom header differs, including empty reports.");
        for (int index = 0; index < document.Pages.Count; index++) {
            PdfReadPage page = document.Pages[index];
            var (width, height) = page.GetPageSize(); Require(width > height && page.GetTextSpans().Count > 0, "Grouped PDF has a blank or incorrectly oriented page.");
            var spans = page.GetTextSpans();
            string footer = string.Concat(spans.Where(span => span.Y < 29.5).OrderBy(span => span.X).Select(span => span.Text));
            JsonElement pageLabels = contract.GetProperty("pageLabels");
            Require(Compact(footer) == Compact(pageLabels.GetProperty("prefix").GetString() + (index + 1) + pageLabels.GetProperty("separator").GetString() + document.Pages.Count), "Explicit host-authored page footer differs.");
            foreach (PdfTextSpan span in spans) Require(span.X >= 29.5 && span.X + span.Advance <= width - 29.5 && span.Y >= 5 && span.Y <= height - 29.5,
                "Grouped text escaped the page's printable bounds: " + JsonSerializer.Serialize(new { span.Text, span.X, span.Y, span.Advance, width, height }) + ".");
        }
        return new { file = Path.GetFileName(path), format = "pdf", rows = rows.Length, groups = groups.Length, pages = document.Pages.Count,
            repaired = false, selectedOrder = true, continuation = true, bytes = new FileInfo(path).Length };
    }

    private static string Compact(string value) => Regex.Replace(value, @"\s+", "");
    private static string Column(int ordinal) { string result = ""; while (ordinal > 0) { ordinal--; result = (char)('A' + ordinal % 26) + result; ordinal /= 26; } return result; }
    private static void Require(bool condition, string message) => WorkbookVerifier.Require(condition, "Grouped report: " + message);
}
