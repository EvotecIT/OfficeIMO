using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetSheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;

/// <summary>Independently check selected inventory, authored formula caches, visible chart references and Excel input rules.</summary>
internal static class EditableAnalysisVerifier {
    private static readonly XNamespace Sheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Chart = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly string[] Statuses = ["Healthy", "Attention", "Offline"];

    internal static object Verify(string path) => VerifyCore(path, path + ".json", edited: false);

    /// <summary>The native copy must contain exactly one Healthy-to-Offline edit and refreshed summary/chart caches.</summary>
    internal static object VerifyNativeEdit(string originalPath, string editedPath) => VerifyCore(editedPath, originalPath + ".json", edited: true);

    private static object VerifyCore(string path, string contractPath, bool edited) {
        using var contractDocument = JsonDocument.Parse(File.ReadAllText(contractPath));
        JsonElement contract = contractDocument.RootElement;
        JsonElement[] expectedRows = contract.GetProperty("rows").EnumerateArray().ToArray();
        string[] headers = contract.GetProperty("headers").EnumerateArray().Select(value => value.GetString()!).ToArray();
        int statusColumn = Array.IndexOf(headers, "Status"), serviceColumn = Array.IndexOf(headers, "Service");
        Require(statusColumn >= 0 && serviceColumn >= 0, "Fixture requires unique selected Status and Service columns.");
        int editedRow = edited ? Array.FindIndex(expectedRows, row => row[statusColumn].GetString() == "Healthy") : -1;
        Require(!edited || editedRow >= 0, "Native edit fixture needs one initial Healthy service.");
        WorkbookVerifier.Verify(path);
        using var package = SpreadsheetDocument.Open(path, false);
        WorkbookPart book = package.WorkbookPart!;
        SpreadsheetSheet[] sheets = book.Workbook!.Sheets!.Elements<SpreadsheetSheet>().ToArray();
        Require(sheets.Select(sheet => sheet.Name!.Value).SequenceEqual(new[] { "Services", "Overview" }), "Expected exactly the two visible report sheets; no duplicate chart-data sheet.");
        Require(sheets.All(sheet => sheet.State is null || sheet.State.Value == SheetStateValues.Visible), "Report data must remain visible.");
        XDocument Read(OpenXmlPart part) { using var stream = part.GetStream(FileMode.Open, FileAccess.Read); return XDocument.Load(stream); }
        var servicesPart = (WorksheetPart)book.GetPartById(sheets[0].Id!.Value!);
        var overviewPart = (WorksheetPart)book.GetPartById(sheets[1].Id!.Value!);
        XDocument services = Read(servicesPart), overview = Read(overviewPart);
        string[] shared = book.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>()
            .Select(item => item.InnerText).ToArray() ?? [];
        Dictionary<string, XElement> Cells(XDocument document) => document.Descendants(Sheet + "c")
            .ToDictionary(cell => cell.Attribute("r")!.Value, StringComparer.Ordinal);
        var serviceCells = Cells(services);
        var overviewCells = Cells(overview);
        string Text(XElement cell) {
            if ((string?)cell.Attribute("t") == "s") return shared[int.Parse(cell.Element(Sheet + "v")!.Value, CultureInfo.InvariantCulture)];
            return (string?)cell.Attribute("t") == "inlineStr" ? string.Concat(cell.Descendants(Sheet + "t").Select(text => text.Value)) : cell.Element(Sheet + "v")?.Value ?? "";
        }
        double Number(XElement cell, string context) {
            Require(cell.Attribute("t") is null || (string?)cell.Attribute("t") == "n", context + " is not numeric.");
            Require(double.TryParse(cell.Element(Sheet + "v")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double result) && double.IsFinite(result), context + " lost its finite numeric cache/value.");
            return result;
        }
        for (int column = 0; column < headers.Length; column++) {
            string coordinate = Column(column + 1) + "2";
            Require(serviceCells.TryGetValue(coordinate, out var heading) && Text(heading) == headers[column], "Selected column heading/order differs at " + coordinate + ".");
        }
        int[] counts = new int[Statuses.Length];
        for (int row = 0; row < expectedRows.Length; row++) for (int column = 0; column < headers.Length; column++) {
            string coordinate = Column(column + 1) + (row + 3).ToString(CultureInfo.InvariantCulture);
            XElement cell = serviceCells.TryGetValue(coordinate, out var captured) ? captured : throw new InvalidDataException("Selected cell missing: " + coordinate);
            Require(cell.Element(Sheet + "f") is null, "Captured data became an authored formula: " + coordinate);
            JsonElement expected = expectedRows[row][column];
            if (column == statusColumn) {
                string status = row == editedRow ? "Offline" : expected.GetString()!;
                Require(Text(cell) == status, "Native status or original raw status differs at " + coordinate);
                int index = Array.IndexOf(Statuses, status); Require(index >= 0, "Unexpected selected status."); counts[index]++;
            } else if (headers[column] == "Checked (UTC)") {
                DateTime instant = DateTimeOffset.Parse(expected.GetString()!, CultureInfo.InvariantCulture).UtcDateTime;
                double serial = (instant - new DateTime(1899, 12, 30, 0, 0, 0, DateTimeKind.Utc)).TotalDays;
                Require(Math.Abs(Number(cell, coordinate) - serial) < 1e-9, "Raw UTC date differs at " + coordinate);
            } else if (expected.ValueKind == JsonValueKind.Number)
                Require(Math.Abs(Number(cell, coordinate) - expected.GetDouble()) < 1e-10, "Raw numeric value differs at " + coordinate);
            else Require(Text(cell) == expected.GetString(), "Raw string/order differs at " + coordinate);
        }
        Require(serviceCells.Keys.Count(address => int.Parse(new string(address.Where(char.IsDigit).ToArray()), CultureInfo.InvariantCulture) is >= 3 &&
            int.Parse(new string(address.Where(char.IsDigit).ToArray()), CultureInfo.InvariantCulture) < expectedRows.Length + 3) == expectedRows.Length * headers.Length, "Captured body gained or lost cells.");
        string statusLetter = Column(statusColumn + 1);
        string dataRange = statusLetter + "3:" + statusLetter + (expectedRows.Length + 2).ToString(CultureInfo.InvariantCulture);
        var validations = services.Descendants(Sheet + "dataValidation").ToArray();
        if (expectedRows.Length == 0) Require(validations.Length == 0 && !servicesPart.TableDefinitionParts.Any(), "Empty inventory invented a data table or invalid validation range.");
        else {
            Require(validations.Length == 1 && (string?)validations[0].Attribute("type") == "list" &&
                (string?)validations[0].Attribute("sqref") == dataRange, "Status list validation is not bound to exactly the selected body rows.");
            Require(validations[0].Element(Sheet + "formula1")?.Value == "\"Healthy,Attention,Offline\"", "Allowed status choices differ.");
            Require((string?)validations[0].Attribute("showErrorMessage") is "1" or "true", "Status validation must expose its native error alert.");
            Require(services.Descendants(Sheet + "conditionalFormatting").Any(rule => (string?)rule.Attribute("sqref") == dataRange), "Editable status highlighting has no live conditional rule.");
            var table = Read(servicesPart.TableDefinitionParts.Single());
            Require((string?)table.Root!.Attribute("name") == "Services", "Analysis formulas lost their named source table.");
        }
        for (int index = 0; index < Statuses.Length; index++) {
            string row = (index + 2).ToString(CultureInfo.InvariantCulture), coordinate = "B" + row;
            Require(Text(overviewCells["A" + row]) == Statuses[index], "Overview category differs.");
            XElement cell = overviewCells[coordinate];
            string sourceRange = "'Services'!$" + statusLetter + "$3:$" + statusLetter + "$" + (expectedRows.Length + 2).ToString(CultureInfo.InvariantCulture);
            string expectedFormula = expectedRows.Length == 0 ? "0" : "COUNTIF(" + sourceRange + ",\"" + Statuses[index] + "\")";
            string actualFormula = cell.Element(Sheet + "f")?.Value ?? "";
            Require(edited ? NativeFormula(actualFormula) == NativeFormula(expectedFormula) : actualFormula == expectedFormula,
                "Authored count formula differs at " + coordinate);
            Require(Number(cell, coordinate) == counts[index], "Formula cache differs at " + coordinate);
        }
        if (!edited) {
            XElement? calculation = Read(book).Root!.Element(Sheet + "calcPr");
            Require(calculation is not null && (string?)calculation.Attribute("calcMode") == "auto" &&
                (string?)calculation.Attribute("fullCalcOnLoad") is "1" or "true", "Authored formulas must request native automatic recalculation.");
        }
        var charts = overviewPart.DrawingsPart?.ChartParts.ToArray() ?? [];
        Require(charts.Length == 1, "Overview needs exactly one native chart.");
        XDocument chart = Read(charts[0]);
        string[] references = chart.Descendants(Chart + "f").Select(formula => formula.Value).ToArray();
        string[] expectedReferences = ["'Overview'!$B$1", "'Overview'!$A$2:$A$4", "'Overview'!$B$2:$B$4"];
        Require(edited ? references.Select(NativeFormula).SequenceEqual(expectedReferences.Select(NativeFormula)) : references.SequenceEqual(expectedReferences),
            "Chart is not bound directly to the visible Overview cells.");
        Require(chart.Descendants(Chart + "numCache").Single().Elements(Chart + "pt")
            .Select(point => double.Parse(point.Element(Chart + "v")!.Value, CultureInfo.InvariantCulture)).SequenceEqual(counts.Select(value => (double)value)), "Native chart numeric cache differs from the calculated summary.");
        Require(chart.Descendants(Chart + "cat").Single().Descendants(Chart + "v").Select(value => value.Value).SequenceEqual(Statuses), "Chart categories differ from visible cells.");
        return new { file = Path.GetFileName(path), profile = contract.GetProperty("profile").GetString(), rows = expectedRows.Length,
            columns = headers.Length, counts, visibleSourceSheets = 2, formulas = 3, charts = 1, statusValidation = expectedRows.Length > 0,
            nativeEdit = edited, openXml = true, readers = 2, bytes = new FileInfo(path).Length };
    }

    private static string Column(int ordinal) {
        string result = "";
        while (ordinal > 0) { ordinal--; result = (char)('A' + ordinal % 26) + result; ordinal /= 26; }
        return result;
    }
    // Native Excel can omit unnecessary local-sheet quotes, retain a leading equals sign,
    // or serialize function names with its compatibility prefix. Coordinates and literals
    // still have to describe the same finite same-workbook expression.
    private static string NativeFormula(string formula) => Regex.Replace(formula.Trim().TrimStart('='), @"\b_xlfn\.", "", RegexOptions.IgnoreCase)
        .Replace("'Services'!", "Services!", StringComparison.OrdinalIgnoreCase).Replace("'Overview'!", "Overview!", StringComparison.OrdinalIgnoreCase)
        .Replace("$", "", StringComparison.Ordinal).ToUpperInvariant();
    private static void Require(bool condition, string message) => WorkbookVerifier.Require(condition, "Editable analysis: " + message);
}
