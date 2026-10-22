using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;

/// <summary>Reads banded inventories independently and checks every selected value, repeated ID, aligned window and page geometry.</summary>
internal static class WideReportVerifier {
    internal static object Verify(string path) {
        using var expected = JsonDocument.Parse(File.ReadAllText(path + ".json")); JsonElement contract = expected.RootElement;
        var document = PdfReadDocument.Open(path, new PdfLoadOptions { IncludeArtifactText = true });
        Require(!document.RepairReport.HasRepairs, "PDF required repair.");
        bool grouped = contract.GetProperty("grouped").GetBoolean();
        string[][] groupPaths = contract.GetProperty("groupPaths").EnumerateArray()
            .Select(path => path.EnumerateArray().Select(value => value.GetString()!).ToArray()).ToArray();
        string[][] rows = contract.GetProperty("rows").EnumerateArray().Select(row => row.EnumerateArray().Select(value => value.GetString()!).ToArray()).ToArray();
        string[] headers = contract.GetProperty("headers").EnumerateArray().Select(value => value.GetString()!).ToArray();
        if (grouped) { rows = rows.Select(row => new[] { "" }.Concat(row).ToArray()).ToArray(); headers = ["Group", .. headers]; }
        double[] widths = contract.GetProperty("widths").EnumerateArray().Select(value => value.GetDouble()).ToArray();
        int[] repeated = contract.GetProperty("repeated").EnumerateArray().Select(value => value.GetInt32()).ToArray();
        var bands = new List<int[]>(); var current = repeated.ToList(); double available = 841.89 - 60, used = repeated.Sum(index => widths[index]);
        for (int column = 0; column < widths.Length; column++) {
            if (repeated.Contains(column)) continue;
            if (used + widths[column] > available + .01) { Require(current.Count > repeated.Length, "Independent band plan cannot fit one data column."); bands.Add(current.ToArray()); current = repeated.ToList(); used = repeated.Sum(index => widths[index]); }
            current.Add(column); used += widths[column];
        }
        if (current.Count > repeated.Length) bands.Add(current.ToArray());
        Require(bands.Count >= 2, "The wide fixture did not require horizontal bands.");
        int idColumn = grouped ? 1 : 0;
        var rowById = rows.ToDictionary(row => row[idColumn], StringComparer.Ordinal);
        var ordinalById = rows.Select((row, index) => (row[idColumn], index)).ToDictionary(pair => pair.Item1, pair => pair.index, StringComparer.Ordinal);
        var seen = rows.ToDictionary(row => row[idColumn], _ => new HashSet<int>(), StringComparer.Ordinal);
        var windows = new Dictionary<string, HashSet<int>>(StringComparer.Ordinal);
        var bandOrders = bands.Select(_ => new List<string>()).ToArray();
        int bodyPages = 0, contextualPages = 0, bodySegments = 0;
        var all = new List<string>();
        for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
            var page = document.Pages[pageIndex]; var spans = page.GetTextSpans(); var (pageWidth, pageHeight) = page.GetPageSize();
            Require(pageWidth > pageHeight, "A wide report unexpectedly changed orientation.");
            string text = Compact(string.Concat(spans.Select(span => span.Text))); all.Add(text);
            Require(spans.Count > 0, "Composed transition produced an empty page.");
            foreach (var span in spans) Require(span.X >= 29.5 && span.X + span.Advance <= pageWidth - 29.5 && span.Y >= 5 && span.Y <= pageHeight - 5,
                "Text escaped printable bounds on page " + (pageIndex + 1) + ": " + span.Text);
            string footer = Compact(string.Concat(spans.Where(span => span.Y < 30).OrderBy(span => span.X).Select(span => span.Text)));
            Require(footer == Compact("Page " + (pageIndex + 1) + " of " + document.Pages.Count), "Global physical page numbering differs on page " + (pageIndex + 1) + ": " + footer);
            var ids = spans.Where(span => Regex.IsMatch(span.Text, @"^S\d{4}$")).ToArray();
            if (ids.Length == 0) continue;
            bodyPages++;
            // Grouped leaves are independent tables and may share a physical page after the preceding table's final band.
            // Keep each contiguous logical window/band separate rather than imposing one band per document page.
            var segments = new List<(int Band, List<PdfTextSpan> Ids)>();
            foreach (var id in ids) {
                Require(rowById.TryGetValue(id.Text, out var expectedRow), "An excluded or invented source ID was emitted.");
                var rowSpans = spans.Where(span => Math.Abs(span.Y - id.Y) < .01).OrderBy(span => span.X).ToArray();
                string[] values = rowSpans.Select(span => span.Text).ToArray();
                int[] matching = Enumerable.Range(0, bands.Count).Where(index => bands[index].Select(column => expectedRow![column]).Where(value => value.Length > 0).SequenceEqual(values)).ToArray();
                Require(matching.Length == 1, "Selected PDF row/cell order differs at " + id.Text + ": " + string.Join(" | ", values));
                int actualBand = matching[0];
                Require(seen[id.Text].Add(actualBand), "A logical source row is duplicated in the same horizontal band.");
                if (segments.Count == 0 || segments[^1].Band != actualBand || ordinalById[segments[^1].Ids[^1].Text] + 1 != ordinalById[id.Text] ||
                    grouped && !groupPaths[ordinalById[segments[^1].Ids[^1].Text]].SequenceEqual(groupPaths[ordinalById[id.Text]]))
                    segments.Add((actualBand, []));
                segments[^1].Ids.Add(id);
                bandOrders[actualBand].Add(id.Text);
                double left = 34; int visible = 0;
                foreach (int column in bands[actualBand]) {
                    if (expectedRow![column].Length > 0) { Require(Math.Abs(rowSpans[visible].X - left) < .1, "Column geometry was scaled below its declared readable width."); visible++; }
                    left += widths[column];
                }
            }
            for (int segmentIndex = 0; segmentIndex < segments.Count; segmentIndex++) {
                var segment = segments[segmentIndex];
                bodySegments++;
                // Only the caption/headings above this segment and below the preceding segment's last data row may satisfy context.
                // A neighboring leaf's caption elsewhere on the same page must not hide a lost ancestor/header.
                double contextTop = segmentIndex == 0 ? pageHeight : segments[segmentIndex - 1].Ids[^1].Y;
                string context = Compact(string.Concat(spans.Where(span => span.Y > segment.Ids[0].Y + .01 && span.Y < contextTop - .01).Select(span => span.Text)));
                foreach (int column in bands[segment.Band]) Require(context.Contains(Compact(headers[column]), StringComparison.Ordinal), "A band lost its captured visible leaf heading.");
                if (contract.GetProperty("structured").GetBoolean()) Require(new[] { "Identity", "Configuration", "Measurements" }.Any(group => context.Contains(group, StringComparison.Ordinal)), "Structured DataTables heading context was lost.");
                string window = string.Join('|', segment.Ids.Select(id => id.Text));
                if (!windows.TryGetValue(window, out var covered)) windows.Add(window, covered = []);
                Require(covered.Add(segment.Band), "The same physical band/window is duplicated.");
                if (grouped) {
                    string[] pathLabels = groupPaths[ordinalById[segment.Ids[0].Text]];
                    foreach (string label in pathLabels) Require(context.Contains(Compact(label), StringComparison.Ordinal), "A horizontal continuation lost an active ancestor: " + label);
                }
            }
            if (grouped) contextualPages++;
        }
        Require(seen.Values.All(coverage => coverage.Count == bands.Count), "A selected ID was omitted from one or more horizontal bands.");
        Require(bandOrders.All(order => order.SequenceEqual(rows.Select(row => row[idColumn]))), "A band's selected record order differs.");
        Require(windows.Count > 1 && windows.Values.All(coverage => coverage.Count == bands.Count), "Logical row windows were not aligned across every band.");
        string full = string.Concat(all);
        foreach (JsonElement forbidden in contract.GetProperty("forbidden").EnumerateArray()) Require(!full.Contains(Compact(forbidden.GetString()!), StringComparison.Ordinal), "Excluded grid data leaked.");
        Require(full.Contains(Compact("Report complete"), StringComparison.Ordinal) && full.Contains(Compact("Selected scope is complete."), StringComparison.Ordinal), "The composed closing section was lost.");
        return new { file = Path.GetFileName(path), rows = rows.Length, columns = headers.Length, bands = bands.Count, pages = document.Pages.Count,
            bodyPages, bodySegments, contextualPages, alignedWindows = windows.Count, selectedOrder = true, readablePointWidths = true, repairs = 0, bytes = new FileInfo(path).Length };
    }
    private static string Compact(string value) => Regex.Replace(value, @"\s+", "");
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidDataException("Wide report: " + message); }
}
