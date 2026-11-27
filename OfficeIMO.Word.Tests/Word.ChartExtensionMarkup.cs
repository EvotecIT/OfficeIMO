using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Chart elements inside extension lists are not chart content: lookups for axes, series and the chart a drawing
/// shows read only the schema positions.
/// </summary>
public sealed class WordChartExtensionMarkupTests : IDisposable {
    private const string ExtensionUri = "{8B2E5C61-3F0A-4D7E-9C1B-2A6F4E8D0C35}";
    private const string ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "officeimo-chart-ext-" + Guid.NewGuid().ToString("N"));

    public WordChartExtensionMarkupTests() => Directory.CreateDirectory(_directory);

    public void Dispose() {
        if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
    }

    [Fact]
    public void AxisIdsInsideChartExtensionsDoNotRaiseTheAxisIdSeed() {
        string path = CreateBarChart("axis-seed.docx");
        RewriteChartPart(path, xml => xml.Replace("</c:chart>",
            "<c:extLst><c:ext uri=\"" + ExtensionUri + "\"><c:axId val=\"4000000000\" /></c:ext></c:extLst></c:chart>"));

        using var document = WordDocument.Load(path);
        WordChart added = document.AddChart("Added");
        added.AddCategories(new[] { "A", "B" });
        added.AddBar("Values", new[] { 1, 2 }, OfficeColor.Black);
        document.Save();

        string addedChart = ReadChartParts(path).Single(xml => xml.Contains("Added", StringComparison.Ordinal));
        foreach (Match axis in Regex.Matches(addedChart, "<c:axId[^>]*val=\"(\\d+)\"", RegexOptions.None, TimeSpan.FromSeconds(5))) {
            Assert.True(uint.Parse(axis.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture) < 4000000000U);
        }
    }

    [Fact]
    public void SeriesInsidePlotAreaExtensionsAreNotEdited() {
        string path = CreateBarChart("series.docx");
        RewriteChartPart(path, xml => xml.Replace("</c:plotArea>",
            "<c:extLst><c:ext uri=\"" + ExtensionUri + "\"><c:barChart><c:ser><c:idx val=\"0\" /><c:order val=\"0\" />" +
            "<c:tx><c:v>Decoy</c:v></c:tx><c:val><c:numLit><c:ptCount val=\"1\" /><c:pt idx=\"0\"><c:v>42</c:v></c:pt></c:numLit></c:val>" +
            "</c:ser></c:barChart></c:ext></c:extLst></c:plotArea>"));

        using (var document = WordDocument.Load(path)) {
            WordChart chart = Assert.Single(document.AllCharts);
            Assert.True(chart.TrySetSeriesValues(0, new[] { 7d, 8d }));
            Assert.True(chart.TrySetSeriesName(0, "Edited"));
            chart.SetDataPointStyle(0, 1, new OfficeChartPointStyle(fillColor: OfficeColor.Parse("#FF0000")));
            document.Save();
        }

        string xml = Assert.Single(ReadChartParts(path));
        int extension = xml.IndexOf("<c:extLst><c:ext uri=\"" + ExtensionUri, StringComparison.Ordinal);
        string content = xml.Substring(0, extension);
        string decoy = xml.Substring(extension);
        Assert.Contains("<c:v>Edited</c:v>", content, StringComparison.Ordinal);
        Assert.Contains("<c:v>8</c:v>", content, StringComparison.Ordinal);
        Assert.Contains("<c:dPt>", content, StringComparison.Ordinal);
        Assert.Contains("<c:v>Decoy</c:v>", decoy, StringComparison.Ordinal);
        Assert.Contains("<c:v>42</c:v>", decoy, StringComparison.Ordinal);
        Assert.DoesNotContain("<c:dPt>", decoy, StringComparison.Ordinal);
    }

    [Fact]
    public void ChartReferenceInsideDrawingExtensionsIsNotTheDrawnChart() {
        string path = CreateBarChart("reference.docx");
        RewritePart(path, "word/document.xml", xml => Regex.Replace(xml, "(<wp:docPr [^>]*?)\\s*/>",
            "$1><a:extLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:ext uri=\"" + ExtensionUri + "\">" +
            "<c:chart xmlns:c=\"" + ChartNs + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdMissing\" />" +
            "</a:ext></a:extLst></wp:docPr>", RegexOptions.None, TimeSpan.FromSeconds(5)));

        using var document = WordDocument.Load(path);
        WordChart chart = Assert.Single(document.AllCharts);
        Assert.Equal("Sales", chart.Title);
        Assert.True(chart.TrySetSeriesValues(0, new[] { 5d, 6d }));
    }

    private string CreateBarChart(string name) {
        string path = Path.Combine(_directory, name);
        using var document = WordDocument.Create(path);
        WordChart chart = document.AddChart("Sales");
        chart.AddCategories(new[] { "North", "South" });
        chart.AddBar("2025", new[] { 1, 2 }, OfficeColor.Parse("#1F4E79"));
        document.Save();
        return path;
    }

    private static List<string> ReadChartParts(string path) {
        using ZipArchive archive = ZipFile.OpenRead(path);
        return archive.Entries.Where(entry => Regex.IsMatch(entry.FullName, "charts/chart\\d+\\.xml$", RegexOptions.None, TimeSpan.FromSeconds(5)))
            .Select(entry => {
                using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
                return reader.ReadToEnd();
            }).ToList();
    }

    private static void RewriteChartPart(string path, Func<string, string> rewrite) =>
        RewritePart(path, name => Regex.IsMatch(name, "charts/chart\\d+\\.xml$", RegexOptions.None, TimeSpan.FromSeconds(5)), rewrite);

    private static void RewritePart(string path, string partName, Func<string, string> rewrite) =>
        RewritePart(path, name => name == partName, rewrite);

    private static void RewritePart(string path, Func<string, bool> match, Func<string, string> rewrite) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.Entries.Single(candidate => match(candidate.FullName));
        string name = entry.FullName;
        string xml;
        using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) xml = reader.ReadToEnd();
        string updated = rewrite(xml);
        Assert.NotEqual(xml, updated);
        entry.Delete();
        using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
        writer.Write(updated);
    }
}
