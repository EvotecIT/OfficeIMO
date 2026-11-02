using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Pdf;

/// <summary>Independent reads of selected hierarchical rows, weighted native summaries and composed report sections.</summary>
internal static class HierarchicalReportVerifier {
    private static readonly XNamespace Sheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Chart = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    internal static object Verify(string path) {
        using var expected = JsonDocument.Parse(File.ReadAllText(path + ".json"));
        return path.EndsWith(".xlsx", StringComparison.Ordinal) ? VerifyWorkbook(path, expected.RootElement) : VerifyPdf(path, expected.RootElement);
    }

    private static object VerifyWorkbook(string path, JsonElement contract) {
        WorkbookVerifier.Verify(path);
        using var package = SpreadsheetDocument.Open(path, false);
        WorkbookPart book = package.WorkbookPart!;
        var sheets = book.Workbook!.Sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().ToArray();
        JsonElement text = contract.GetProperty("text");
        Require(sheets.Length == 3 && sheets.Select(sheet => sheet.Name!.Value!).SequenceEqual(new[] { Text(text, "context"), Text(text, "detail"), Text(text, "overview") }) &&
            sheets.All(sheet => sheet.State is null || sheet.State.Value == SheetStateValues.Visible), "Context, grouped detail and overview sheets differ.");
        WorksheetPart Part(int index) => (WorksheetPart)book.GetPartById(sheets[index].Id!.Value!);
        string[] shared = book.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().Select(item => item.InnerText).ToArray() ?? [];
        var context = new SheetCells(Part(0), shared);
        var detail = new SheetCells(Part(1), shared);
        var overview = new SheetCells(Part(2), shared);
        Require(context.AllText.Contains(Text(text, "title"), StringComparison.Ordinal) && context.AllText.Contains(Text(text, "scope"), StringComparison.Ordinal), "Context sheet lost authored title/scope.");
        Require(!Part(1).TableDefinitionParts.Any() && !detail.Xml.Descendants(Sheet + "autoFilter").Any(), "Grouped summaries entered a native data table/autofilter.");
        string[] keys = Strings(contract.GetProperty("keys")), headers = Strings(contract.GetProperty("headers"));
        JsonElement[] rows = contract.GetProperty("rows").EnumerateArray().ToArray(), events = contract.GetProperty("events").EnumerateArray().ToArray();
        int depth = contract.GetProperty("groupKeys").GetArrayLength();
        Require(detail.Text(2, 1) == Text(text, "group"), "Leading grouped heading differs.");
        for (int column = 0; column < headers.Length; column++) Require(detail.Text(2, column + 2) == headers[column], "Current visible heading/order differs.");
        int output = 3, accepted = 0, firstDetail = 0, lastDetail = 0, formulas = 0;
        var active = new Stack<GroupRange>(); var detailRows = new HashSet<int>();
        foreach (JsonElement item in events) {
            switch (Text(item, "kind")) {
                case "groupStart":
                    Require(detail.Text(output, 1) == Text(item, "label"), "Typed group opening order/label differs.");
                    Require(Outline(detail, output) == item.GetProperty("level").GetInt32(), "Group heading outline depth differs.");
                    active.Push(new GroupRange()); break;
                case "row":
                    int index = item.GetProperty("index").GetInt32(); Require(index == accepted, "Expected selected source order is incomplete.");
                    Require(Outline(detail, output) == depth, "Detail outline lost an ancestor level.");
                    for (int column = 0; column < keys.Length; column++) Value(detail.Cell(output, column + 2), detail, rows[index][column], keys[column]);
                    firstDetail = firstDetail == 0 ? output : firstDetail; lastDetail = output; detailRows.Add(output);
                    foreach (GroupRange group in active) { if (group.First == 0) group.First = output; group.Last = output; }
                    accepted++; break;
                case "groupEnd":
                    GroupRange ending = active.Pop();
                    Require(detail.Text(output, 1) == Text(item, "summaryLabel"), "Group subtotal count/label differs.");
                    formulas += Summary(detail, output, keys, contract.GetProperty("operations"), item.GetProperty("totals"), ending.First, ending.Last);
                    break;
                default: throw new InvalidDataException("Unexpected independent hierarchy event.");
            }
            output++;
        }
        Require(active.Count == 0 && accepted == rows.Length, "Selected hierarchy rows or closing groups differ.");
        Require(detail.Text(output, 1) == Text(contract, "grandLabel"), "Grand total count/label differs.");
        formulas += Summary(detail, output, keys, contract.GetProperty("operations"), contract.GetProperty("totals"), firstDetail, lastDetail);
        Require(detail.Xml.Descendants(Sheet + "row").Count() == output && detail.Xml.Descendants(Sheet + "f").Count() == formulas, "Unexpected summary/body rows or formulas.");
        Require(detail.Xml.Descendants(Sheet + "row").All(row => row.Elements(Sheet + "c").Count() <= keys.Length + 1), "Hidden grouping fields became duplicated output columns.");
        Require(!detail.AllText.Contains("UNSELECTED-", StringComparison.Ordinal) && !detail.AllText.Contains("INTERNAL-ONLY", StringComparison.Ordinal), "Excluded grid data leaked.");
        if (rows.Length > 0 && keys.Contains("status", StringComparer.Ordinal))
            Require(detail.Xml.Descendants(Sheet + "conditionalFormatting").Any() && detail.Xml.Descendants(Sheet + "dataValidation").Any(), "Selected status details lost native formatting/editing rules.");
        foreach (XElement rule in detail.Xml.Descendants(Sheet + "conditionalFormatting").Concat(detail.Xml.Descendants(Sheet + "dataValidation"))) {
            string references = (string?)rule.Attribute("sqref") ?? throw new InvalidDataException("A detail rule has no finite range.");
            foreach (string reference in references.Split(' ', StringSplitOptions.RemoveEmptyEntries)) {
                var match = Regex.Match(reference, @"^[A-Z]+(\d+)(?::[A-Z]+(\d+))?$"); Require(match.Success, "Detail rules require finite cell ranges.");
                int first = int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture), last = match.Groups[2].Success ? int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture) : first;
                Require(Enumerable.Range(first, last - first + 1).All(detailRows.Contains), "Editing/format rules include generated hierarchy rows.");
            }
        }
        JsonElement[] parents = events.Where(item => Text(item, "kind") == "groupEnd" && item.GetProperty("level").GetInt32() == 0).ToArray();
        for (int index = 0; index < parents.Length; index++) {
            Require(overview.Text(index + 2, 1) == Text(parents[index], "label"), "Overview category order differs.");
            Number(overview.Cell(index + 2, 2), parents[index].GetProperty("rows").GetDouble());
        }
        var charts = Part(2).DrawingsPart?.ChartParts.ToArray() ?? [];
        Require(charts.Length == (parents.Length == 0 ? 0 : 1), "Native overview chart count differs.");
        if (charts.Length != 0) {
            using var stream = charts[0].GetStream(FileMode.Open, FileAccess.Read); var chart = XDocument.Load(stream);
            Require(chart.Descendants(Chart + "numCache").Single().Elements(Chart + "pt").Select(point => double.Parse(point.Element(Chart + "v")!.Value, CultureInfo.InvariantCulture))
                .SequenceEqual(parents.Select(parent => parent.GetProperty("rows").GetDouble())), "Overview chart cache differs from independently counted selected records.");
        }
        return new { file = Path.GetFileName(path), format = "xlsx", rows = accepted, levels = depth, groups = events.Count(item => Text(item, "kind") == "groupEnd"),
            formulas, visibleSheets = 3, nativeCharts = charts.Length, openXml = true, readers = 2, detailOnlyEditing = true, bytes = new FileInfo(path).Length };
    }

    private static int Outline(SheetCells sheet, int row) => (int?)sheet.Xml.Descendants(Sheet + "row").Single(item => (int)item.Attribute("r")! == row).Attribute("outlineLevel") ?? 0;

    private static int Summary(SheetCells sheet, int row, string[] keys, JsonElement operations, JsonElement totals, int first, int last) {
        int formulas = 0;
        for (int column = 0; column < keys.Length; column++) if (operations.TryGetProperty(keys[column], out var operation)) {
            XElement? cell = sheet.Cell(row, column + 2); JsonElement expected = totals[column];
            if (first == 0) { if (expected.ValueKind == JsonValueKind.Null) Require(sheet.Content(cell) == "", "Empty numeric aggregate gained a value."); else Number(cell, expected.GetDouble()); continue; }
            string range = Column(column + 2) + first.ToString(CultureInfo.InvariantCulture) + ":" + Column(column + 2) + last.ToString(CultureInfo.InvariantCulture);
            string code = operation.GetString() switch { "average" => "1", "count" => "2", "max" => "4", "min" => "5", "sum" => "9", _ => throw new InvalidDataException("Unknown independent aggregate.") };
            string expression = "SUBTOTAL(" + code + "," + range + ")";
            if (operation.GetString() is "average" or "min" or "max") expression = "IF(SUBTOTAL(2," + range + ")=0,\"\"," + expression + ")";
            Require(cell?.Element(Sheet + "f")?.Value == expression, "A native subtotal includes the wrong range/function or loses its empty guard.");
            if (expected.ValueKind == JsonValueKind.Null) Require(cell?.Element(Sheet + "v")?.Value is null or "", "Empty aggregate cache gained a number.");
            else Number(cell, expected.GetDouble());
            formulas++;
        }
        return formulas;
    }

    private static void Value(XElement? cell, SheetCells sheet, JsonElement expected, string key) {
        Require(cell?.Element(Sheet + "f") is null, "A selected literal value became a formula.");
        if (expected.ValueKind == JsonValueKind.Null) Require(sheet.Content(cell) == "", "A selected null became zero/text.");
        else if (expected.ValueKind == JsonValueKind.Number) Number(cell, expected.GetDouble());
        else if (key == "checked") Number(cell, (DateTimeOffset.Parse(expected.GetString()!, CultureInfo.InvariantCulture).UtcDateTime - new DateTime(1899, 12, 30)).TotalDays);
        else Require(sheet.Content(cell) == expected.GetString(), "Selected raw text/order differs.");
    }

    private static void Number(XElement? cell, double expected) {
        Require((string?)cell?.Attribute("t") is null or "n", "A numeric/date value or cache became text.");
        Require(double.TryParse(cell?.Element(Sheet + "v")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double actual) && double.IsFinite(actual) && Math.Abs(actual - expected) < 1e-9,
            "Selected numeric/date value or independent aggregate cache differs.");
    }

    private static object VerifyPdf(string path, JsonElement contract) {
        using var document = PdfDocument.Load(path);
        using var painted = PdfDocument.Load(path, new PdfLoadOptions { IncludeArtifactText = true });
        Require(!document.RepairReport.HasRepairs, "Composed PDF required repair.");
        JsonElement text = contract.GetProperty("text"), pageLabels = text.GetProperty("page");
        JsonElement[] rows = contract.GetProperty("rows").EnumerateArray().ToArray(), events = contract.GetProperty("events").EnumerateArray().ToArray();
        string[] keys = Strings(contract.GetProperty("keys")), headers = Strings(contract.GetProperty("headers"));
        int service = Array.IndexOf(keys, "service"); Require(service >= 0, "The selected-ID contract is missing.");
        string[] pageTexts = document.Pages.Select(page => string.Concat(page.GetTextSpans().Select(span => span.Text))).ToArray();
        string all = string.Concat(pageTexts), compact = Compact(all);
        Require(Regex.Matches(all, @"WEB-\d{6}").Select(match => match.Value).SequenceEqual(rows.Select(row => row[service].GetString()!)), "PDF selected IDs were lost, duplicated or reordered.");
        Require(!all.Contains("UNSELECTED-", StringComparison.Ordinal) && !all.Contains("INTERNAL-ONLY", StringComparison.Ordinal), "Excluded grid data leaked into PDF.");
        Require(document.Outlines.Count == 1, "The composed report has no unique root destination.");
        PdfOutlineItem[] outlines = [document.Outlines[0], .. document.Outlines[0].Children];
        Require(outlines.Select(item => item.Title).SequenceEqual(new[] { Text(text, "title"), Text(text, "detail"), Text(text, "next") }), "Summary, inventory and recommendation destinations differ.");
        Require(outlines.All(item => item.PageNumber is int target && target >= 1 && target <= document.Pages.Count && item.DestinationTop.HasValue), "A report destination is unresolved.");
        int inventory = outlines[1].PageNumber!.Value - 1, next = outlines[2].PageNumber!.Value - 1;
        Require(inventory > 0 && next > inventory, "Composed report sections were lost or overlap.");
        var culture = CultureInfo.GetCultureInfo(Text(contract, "language"));
        string Format(JsonElement value, string key) => value.ValueKind == JsonValueKind.Null ? "" : value.ValueKind == JsonValueKind.Number ? value.GetDouble().ToString("#,0.###", culture)
            : key == "checked" ? DateTimeOffset.Parse(value.GetString()!, CultureInfo.InvariantCulture).UtcDateTime.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) : value.GetString()!;
        string AggregateText(JsonElement totals) => string.Concat(totals.EnumerateArray().Select(value => Format(value, "")));
        foreach (JsonElement ending in events.Where(item => Text(item, "kind") == "groupEnd"))
            Require(compact.Contains(Compact(Text(ending, "summaryLabel") + AggregateText(ending.GetProperty("totals"))), StringComparison.Ordinal), "PDF subtotal label/count or independently weighted metrics differ.");
        Require(compact.Contains(Compact(Text(contract, "grandLabel") + AggregateText(contract.GetProperty("totals"))), StringComparison.Ordinal), "PDF grand total differs from selected raw records.");
        Require(compact.Contains(Compact(Text(text, "scope")), StringComparison.Ordinal) && compact.Contains(Compact(Text(text, "advice")), StringComparison.Ordinal), "Summary scope or recommendations were lost.");
        if (contract.TryGetProperty("figureCaption", out var caption)) Require(compact.Contains(Compact(caption.GetString()!), StringComparison.Ordinal), "Canonical ChartForgeX figure caption was lost.");
        var paths = new Dictionary<int, string[]>(); var active = new List<string>();
        foreach (JsonElement item in events) {
            switch (Text(item, "kind")) {
                case "groupStart": active.Add(Text(item, "label")); break;
                case "groupEnd": active.RemoveAt(active.Count - 1); break;
                case "row": paths.Add(item.GetProperty("index").GetInt32(), active.ToArray()); break;
            }
        }
        int accepted = 0, previousPageLast = -1, continued = 0;
        for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
            var page = document.Pages[pageIndex]; var (width, height) = page.GetPageSize(); var spans = page.GetTextSpans();
            bool appendix = pageIndex >= inventory && pageIndex < next;
            Require(appendix ? width > height : height > width, "Composed portrait/landscape section geometry differs.");
            Require(spans.Count > 0, "Composed transition produced a blank page.");
            var ids = spans.Where(span => Regex.IsMatch(span.Text, @"^WEB-\d{6}$")).ToArray();
            if (appendix) {
                string paintedText = Compact(string.Concat(painted.Pages[pageIndex].GetTextSpans().Select(span => span.Text)));
                foreach (string header in headers) Require(paintedText.Contains(Compact(header), StringComparison.Ordinal), "A grouped continuation lost a visible leaf heading.");
                if (ids.Length > 0 && previousPageLast >= 0) {
                    string[] prior = paths[previousPageLast], current = paths[accepted];
                    int common = 0; while (common < Math.Min(prior.Length, current.Length) && prior[common] == current[common]) common++;
                    if (common > 0) {
                        foreach (string ancestor in current.Take(common)) Require(paintedText.Contains(Compact(ancestor), StringComparison.Ordinal), "A continued page lost an active ancestor context.");
                        if (common == current.Length && prior.Length == current.Length)
                            Require(paintedText.Contains(Compact(Text(text, "continued")), StringComparison.Ordinal), "A continued leaf group lost its explicit continuation label.");
                        continued++;
                    }
                }
                foreach (PdfTextSpan id in ids) {
                    string[] expected = rows[accepted].EnumerateArray().Select((value, column) => Format(value, keys[column])).Where(value => value.Length > 0).ToArray();
                    string[] actual = spans.Where(span => Math.Abs(span.Y - id.Y) < .01).OrderBy(span => span.X).Select(span => span.Text).ToArray();
                    Require(actual.SequenceEqual(expected), "PDF selected visible cells differ at source row " + accepted + ": " + string.Join(" | ", actual));
                    accepted++;
                }
                if (ids.Length > 0) previousPageLast = accepted - 1;
            }
            string sectionTitle = pageIndex < inventory ? Text(text, "title") : pageIndex < next ? Text(text, "detail") : Text(text, "next");
            // Untagged documents expose the authored running decorations through the ordinary text reader.
            foreach (PdfTextSpan span in spans) Require(span.X >= 29.5 && span.X + span.Advance <= width - 29.5 && span.Y >= 5 && span.Y <= height - 5 &&
                (span.Y <= height - 29.5 || span.Text == sectionTitle),
                "Analytical text escaped its body or authored decoration bounds: " + JsonSerializer.Serialize(new { page = pageIndex + 1, span.Text, span.X, span.Y, span.Advance, width, height }) + ".");
            var footers = painted.Pages[pageIndex].GetTextSpans().Where(span => span.Y < 30).ToArray();
            string footer = Compact(string.Concat(footers.OrderBy(span => span.X).Select(span => span.Text)));
            Require(footer == Compact(Text(pageLabels, "prefix") + (pageIndex + 1) + Text(pageLabels, "separator") + document.Pages.Count + Text(pageLabels, "suffix")),
                "Composed global page numbering differs.");
        }
        Require(accepted == rows.Length && (rows.Length < 50 || continued > 0), "Selected rows or representative ancestor continuation proof is incomplete.");
        return new { file = Path.GetFileName(path), format = "pdf", rows = accepted, levels = contract.GetProperty("groupKeys").GetArrayLength(), pages = document.Pages.Count,
            continuedPages = continued, composedDestinations = outlines.Length, selectedOrder = true, independentMetrics = true, repairs = 0, bytes = new FileInfo(path).Length };
    }

    private sealed class GroupRange { internal int First; internal int Last; }
    private sealed class SheetCells {
        private readonly Dictionary<string, XElement> cells;
        private readonly string[] shared;
        internal XDocument Xml { get; }
        internal SheetCells(WorksheetPart part, string[] strings) {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read); Xml = XDocument.Load(stream); shared = strings;
            cells = Xml.Descendants(Sheet + "c").ToDictionary(cell => cell.Attribute("r")!.Value, StringComparer.Ordinal);
        }
        internal XElement? Cell(int row, int column) => cells.GetValueOrDefault(Column(column) + row.ToString(CultureInfo.InvariantCulture));
        internal string Content(XElement? cell) => cell is null ? "" : (string?)cell.Attribute("t") == "s" ? shared[int.Parse(cell.Element(Sheet + "v")!.Value, CultureInfo.InvariantCulture)]
            : (string?)cell.Attribute("t") == "inlineStr" ? string.Concat(cell.Descendants(Sheet + "t").Select(value => value.Value)) : cell.Element(Sheet + "v")?.Value ?? "";
        internal string Text(int row, int column) => Content(Cell(row, column));
        internal string AllText => string.Concat(cells.Values.Select(Content));
    }
    private static string Text(JsonElement element, string name) => element.GetProperty(name).GetString()!;
    private static string[] Strings(JsonElement array) => array.EnumerateArray().Select(value => value.GetString()!).ToArray();
    private static string Compact(string value) => Regex.Replace(value, @"\s+", "");
    private static string Column(int ordinal) { string result = ""; while (ordinal > 0) { ordinal--; result = (char)('A' + ordinal % 26) + result; ordinal /= 26; } return result; }
    private static void Require(bool condition, string message) => WorkbookVerifier.Require(condition, "Analytical report: " + message);
}
