using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;

/// <summary>Independent semantic checks for downloads made by the report workbench UI.</summary>
internal static class ReportEngineVerifier {
    private static readonly string[] Headers = ["Service", "Site", "Status", "Latency (ms)", "Availability", "Checked (UTC)"];
    private static readonly string[] Sites = ["Warsaw", "Łódź", "Praha", "Berlin"];
    private static readonly XNamespace Sheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Chart = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    internal static object Verify(string path, int count) => Path.GetExtension(path) switch {
        ".csv" => VerifyCsv(path, count), ".xlsx" => VerifyXlsx(path, count), ".pdf" => VerifyPdf(path, count),
        _ => throw new InvalidDataException("Unsupported workbench artifact.")
    };

    private static string Service(int index) => "WEB-" + (index + 1).ToString("D6", CultureInfo.InvariantCulture);
    private static string Status(int index) => index % 37 == 0 ? "Offline" : index % 11 == 0 ? "Attention" : "Healthy";
    private static double Latency(int index) => index % 97 / 10.0 + 4.75;
    private static double Availability(int index) => 0.995 - index % 17 / 10000.0;
    private static DateTime Checked(int index) => new(2026, 10, 19, 8, index % 60, 0, DateTimeKind.Utc);
    private static void Require(bool condition, string message) => WorkbookVerifier.Require(condition, "Report workbench: " + message);
    private static void Number(object? actual, double expected, string context) => Require(actual is not string && actual is not null &&
        Math.Abs(Convert.ToDouble(actual, CultureInfo.InvariantCulture) - expected) < 1e-10, context + " lost its numeric type/value.");

    private static object VerifyCsv(string path, int count) {
        Require(File.ReadAllBytes(path).Take(3).SequenceEqual(new byte[] { 239, 187, 191 }), "CSV lost its UTF-8 BOM.");
        using var reader = CsvDocument.OpenDataReader(path);
        Require(reader.FieldCount == Headers.Length && Enumerable.Range(0, Headers.Length).Select(reader.GetName).SequenceEqual(Headers), "CSV headers differ.");
        int row = 0;
        while (reader.Read()) {
            Require(row < count && reader.GetString(0) == Service(row) && reader.GetString(1) == Sites[row % 4] && reader.GetString(2) == Status(row), "CSV row order/text differs.");
            Require(Math.Abs(double.Parse(reader.GetString(3), CultureInfo.InvariantCulture) - Latency(row)) < 1e-10 &&
                Math.Abs(double.Parse(reader.GetString(4), CultureInfo.InvariantCulture) - Availability(row)) < 1e-10, "CSV numeric precision differs.");
            Require(DateTimeOffset.Parse(reader.GetString(5), CultureInfo.InvariantCulture).UtcDateTime == Checked(row) && reader.GetString(5).EndsWith("Z", StringComparison.Ordinal), "CSV UTC timestamp differs.");
            row++;
        }
        Require(row == count, "CSV row count differs.");
        return new { file = Path.GetFileName(path), rows = row, columns = Headers.Length, bytes = new FileInfo(path).Length, reader = "OfficeIMO.CSV" };
    }

    private static object VerifyXlsx(string path, int count) {
        WorkbookVerifier.Verify(path);
        using (var reader = ExcelDocument.OpenReader(path, new ExcelReaderOptions { SheetName = "Services", A1Range = "A2:F" + (count + 2), HasHeaderRow = true })) {
            Require(reader.FieldCount == Headers.Length && Enumerable.Range(0, Headers.Length).Select(reader.GetName).SequenceEqual(Headers), "XLSX headings differ.");
            int row = 0;
            while (reader.Read()) {
                Require(row < count && Equals(reader.GetValue(0), Service(row)) && Equals(reader.GetValue(1), Sites[row % 4]) && Equals(reader.GetValue(2), Status(row)), "XLSX row order/text differs.");
                Number(reader.GetValue(3), Latency(row), "XLSX latency"); Number(reader.GetValue(4), Availability(row), "XLSX availability");
                Require(reader.GetValue(5) is DateTime date && Math.Abs((date - Checked(row)).TotalMilliseconds) < 1, "XLSX UTC clock/date type differs.");
                row++;
            }
            Require(row == count, "XLSX row count differs.");
        }
        using var zip = ZipFile.OpenRead(path);
        XDocument Read(string name) { using var stream = zip.GetEntry(name)?.Open() ?? throw new InvalidDataException("Missing " + name); return XDocument.Load(stream); }
        var sheets = Read("xl/workbook.xml").Descendants(Sheet + "sheet").ToArray();
        Require(sheets.Length == 2 && (string?)sheets[0].Attribute("name") == "Services" && (string?)sheets[1].Attribute("name") == "Overview" && sheets.All(sheet => sheet.Attribute("state") is null), "XLSX report/chart sheet identities differ.");
        var service = Read("xl/worksheets/sheet1.xml");
        Require((string?)service.Descendants(Sheet + "pane").Single().Attribute("topLeftCell") == "B3", "XLSX frozen headings differ.");
        if (count > 0) {
            int footer = count + 3;
            foreach ((string column, double expected) in new[] { ("D", Enumerable.Range(0, count).Average(Latency)), ("E", Enumerable.Range(0, count).Average(Availability)) }) {
                var cell = service.Descendants(Sheet + "c").Single(c => (string?)c.Attribute("r") == column + footer);
                string range = $"{column}3:{column}{count + 2}";
                Require((string?)cell.Element(Sheet + "f") == $"IF(SUBTOTAL(102,{range})=0,\"\",SUBTOTAL(101,{range}))", "XLSX guarded average formula range differs.");
                Require(Math.Abs(double.Parse(cell.Element(Sheet + "v")!.Value, CultureInfo.InvariantCulture) - expected) < 1e-10, "XLSX average cache differs.");
            }
            Require((string?)Read("xl/tables/table1.xml").Root!.Attribute("ref") == "A2:F" + footer, "XLSX native table bounds differ.");
        } else Require(!zip.Entries.Any(e => e.FullName.StartsWith("xl/tables/", StringComparison.Ordinal)), "Empty XLSX invented a data table.");
        var chart = Read("xl/charts/chart1.xml");
        Require(chart.Descendants(Chart + "barChart").Count() == 1 && (string?)chart.Descendants(Chart + "barDir").Single().Attribute("val") == "col", "Native clustered column chart missing.");
        Require((string?)chart.Descendants(Chart + "plotVisOnly").Single().Attribute("val") == "1", "Visible source chart filtering policy differs.");
        int[] counts = new[] { "Healthy", "Attention", "Offline" }.Select(status => Enumerable.Range(0, count).Count(i => Status(i) == status)).ToArray();
        Require(chart.Descendants(Chart + "numCache").Single().Elements(Chart + "pt").Select(p => int.Parse(p.Element(Chart + "v")!.Value, CultureInfo.InvariantCulture)).SequenceEqual(counts), "Chart numeric cache differs.");
        var source = Read("xl/worksheets/sheet2.xml");
        string Cell(string reference) => source.Descendants(Sheet + "c").Single(c => (string?)c.Attribute("r") == reference).Element(Sheet + "v")?.Value ??
            string.Concat(source.Descendants(Sheet + "c").Single(c => (string?)c.Attribute("r") == reference).Descendants(Sheet + "t").Select(t => t.Value));
        Require(Enumerable.Range(2, 3).Select(i => int.Parse(Cell("B" + i), CultureInfo.InvariantCulture)).SequenceEqual(counts), "Chart backing source values differ.");
        for (int index = 0; index < 3; index++) {
            string status = new[] { "Healthy", "Attention", "Offline" }[index];
            XElement summary = source.Descendants(Sheet + "c").Single(cell => (string?)cell.Attribute("r") == "B" + (index + 2));
            string expected = count == 0 ? "0" : "COUNTIF('Services'!$C$3:$C$" + (count + 2) + ",\"" + status + "\")";
            Require(summary.Element(Sheet + "f")?.Value == expected, "Live service summary formula differs.");
        }
        if (count > 0) Require((string?)service.Descendants(Sheet + "dataValidation").Single().Attribute("sqref") == "C3:C" + (count + 2), "Status editing validation range differs.");
        string prefix = "'" + sheets[1].Attribute("name")!.Value.Replace("'", "''", StringComparison.Ordinal) + "'!";
        Require(chart.Descendants(Chart + "f").Select(f => f.Value).SequenceEqual(new[] { prefix + "$B$1", prefix + "$A$2:$A$4", prefix + "$B$2:$B$4" }), "Chart source formulas differ.");
        Require(chart.Descendants(Chart + "cat").Single().Descendants(Chart + "v").Select(v => v.Value).SequenceEqual(new[] { "Healthy", "Attention", "Offline" }) &&
            Enumerable.Range(2, 3).Select(i => Cell("A" + i)).SequenceEqual(new[] { "Healthy", "Attention", "Offline" }), "Chart categories differ from source cells.");
        return new { file = Path.GetFileName(path), rows = count, columns = Headers.Length, sheets = 2, charts = 1, bytes = new FileInfo(path).Length, readers = 2, openXml = true, formulasAndCaches = true, typedDates = "UTC clock" };
    }

    private static object VerifyPdf(string path, int count) {
        PdfReadDocument document = PdfReadDocument.Open(path);
        Require(!document.RepairReport.HasRepairs, "PDF required repairs.");
        string[] spans = document.Pages.SelectMany(page => page.GetTextSpans()).Select(span => span.Text).ToArray();
        string text = string.Concat(spans);
        string normalized = Regex.Replace(string.Join(" ", spans), @"\s+", " ");
        int[] ids = Regex.Matches(text, @"WEB-(\d{6})").Select(match => int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture)).ToArray();
        Require(ids.SequenceEqual(Enumerable.Range(1, count)), "PDF service IDs were lost, duplicated or reordered.");
        foreach (string heading in Headers) Require(normalized.Contains(heading, StringComparison.Ordinal), "PDF heading missing: " + heading);
        Require(normalized.Contains("Service health report", StringComparison.Ordinal) && normalized.Contains(count.ToString(CultureInfo.InvariantCulture) + " services.", StringComparison.Ordinal), "PDF title/summary differs: " + normalized[..Math.Min(120, normalized.Length)] + " ... " + normalized[Math.Max(0, normalized.Length - 160)..]);
        if (count > 0) {
            Require(text.Contains("Łódź", StringComparison.Ordinal) && text.Contains("2026-10-19 08:00", StringComparison.Ordinal) && text.Contains("2026-10-19 08:39", StringComparison.Ordinal), "PDF Unicode or UTC values differ.");
            Require(text.Contains("4.9", StringComparison.Ordinal) && text.Contains("99.50%", StringComparison.Ordinal), "PDF displayed numeric values differ.");
        }
        return new { file = Path.GetFileName(path), rows = count, pages = document.Pages.Count, bytes = new FileInfo(path).Length, repairs = 0, reader = "OfficeIMO.Pdf" };
    }
}
