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
    private const string McNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string Wp14Ns = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
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

        string addedChart = ReadChartParts(path).Single(xml => xml.IndexOf("Added", StringComparison.Ordinal) >= 0);
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

    [Fact]
    public void AxisIdsThatAreNotNumbersDoNotStopTheDocumentLoading() {
        string path = CreateBarChart("axis-text.docx");
        RewriteChartPart(path, xml => xml.Replace("</c:chart>", "<c:axId val=\"not-a-number\" /></c:chart>"));

        using var document = WordDocument.Load(path);
        Assert.Equal("Sales", Assert.Single(document.AllCharts).Title);
    }

    [Fact]
    public void TitleRunsInsideVendorMarkupAreNotTitleText() {
        string path = CreateBarChart("vendor-title.docx");
        RewriteChartPart(path, xml => ReplaceFirst(xml, "<a:r>",
            "<v:meta xmlns:v=\"urn:example:vendor\"><a:r><a:t>Meta</a:t></a:r></v:meta><a:r>"));

        using (var document = WordDocument.Load(path)) {
            WordChart chart = Assert.Single(document.AllCharts);
            Assert.Equal("Sales", chart.Title);
            chart.SetTitle("Renamed");
            document.Save();
        }

        using (var document = WordDocument.Load(path)) {
            Assert.Equal("Renamed", Assert.Single(document.AllCharts).Title);
        }
        string xml = Assert.Single(ReadChartParts(path));
        Assert.Contains("<a:t>Renamed</a:t>", xml, StringComparison.Ordinal);
        Assert.Contains("<a:t>Meta</a:t>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void ChartFrameInsideVendorDrawingMarkupIsNotTheDrawnChart() {
        string path = CreateBarChart("vendor-reference.docx");
        RewritePart(path, "word/document.xml", xml => Regex.Replace(xml, "(<wp:docPr [^>]*?)\\s*/>",
            "$1><v:payload xmlns:v=\"urn:example:vendor\"><wp:inline><a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
            "<a:graphicData uri=\"" + ChartNs + "\">" +
            "<c:chart xmlns:c=\"" + ChartNs + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdMissing\" />" +
            "</a:graphicData></a:graphic></wp:inline></v:payload></wp:docPr>",
            RegexOptions.None, TimeSpan.FromSeconds(5)));

        using var document = WordDocument.Load(path);
        WordChart chart = Assert.Single(document.AllCharts);
        Assert.Equal("Sales", chart.Title);
        Assert.True(chart.TrySetSeriesName(0, "Edited"));
    }

    [Fact]
    public void ChartFrameInsideAlternateContentIsTheDrawnChart() {
        string path = CreateBarChart("alternate-content.docx");
        RewritePart(path, "word/document.xml", xml => Regex.Replace(xml, "(<wp:(inline|anchor)\\b.*?</wp:\\2>)",
            "<mc:AlternateContent xmlns:mc=\"" + McNs + "\"><mc:Choice Requires=\"wp14\" xmlns:wp14=\"" + Wp14Ns + "\">$1</mc:Choice></mc:AlternateContent>",
            RegexOptions.Singleline, TimeSpan.FromSeconds(5)));

        using (var document = WordDocument.Load(path)) {
            WordChart chart = Assert.Single(document.AllCharts);
            Assert.Equal("Sales", chart.Title);
            Assert.True(chart.TrySetSeriesName(0, "Edited"));
            chart.Name = "Sales chart";
            chart.AltText = "Sales by region";
            document.Save();
        }

        using (var document = WordDocument.Load(path)) {
            WordChart chart = Assert.Single(document.AllCharts);
            Assert.Equal("Sales chart", chart.Name);
            Assert.Equal("Sales by region", chart.AltText);
        }
        Assert.Contains("<c:v>Edited</c:v>", Assert.Single(ReadChartParts(path)), StringComparison.Ordinal);
    }

    [Fact]
    public void ChartReferenceInsideAlternateContentIsTheDrawnChart() {
        string path = CreateBarChart("alternate-reference.docx");
        RewritePart(path, "word/document.xml", xml => Regex.Replace(xml, "(<c:chart\\b[^>]*/>)",
            "<mc:AlternateContent xmlns:mc=\"" + McNs + "\"><mc:Choice Requires=\"wp14\" xmlns:wp14=\"" + Wp14Ns + "\">$1</mc:Choice></mc:AlternateContent>",
            RegexOptions.None, TimeSpan.FromSeconds(5)));

        using var document = WordDocument.Load(path);
        WordChart chart = Assert.Single(document.AllCharts);
        Assert.Equal("Sales", chart.Title);
        Assert.True(chart.TrySetSeriesName(0, "Edited"));
    }

    [Theory]
    [InlineData("v", "rIdMissing", "{0}")]
    [InlineData("wp14", "{0}", "rIdMissing")]
    public void AlternateContentUsesTheBranchWordSelects(string requires, string choiceId, string fallbackId) {
        string path = CreateBarChart("alternate-" + requires + ".docx");
        RewritePart(path, "word/document.xml", xml => Regex.Replace(xml, "<c:chart\\b([^>]*?)(\\w+):id=\"([^\"]+)\"([^>]*)/>", match => {
            string Reference(string id) => "<c:chart" + match.Groups[1].Value + match.Groups[2].Value + ":id=\"" +
                string.Format(System.Globalization.CultureInfo.InvariantCulture, id, match.Groups[3].Value) + "\"" + match.Groups[4].Value + "/>";
            return "<mc:AlternateContent xmlns:mc=\"" + McNs + "\" xmlns:v=\"urn:example:vendor\" xmlns:wp14=\"" + Wp14Ns + "\"><mc:Choice Requires=\"" + requires + "\">" + Reference(choiceId) +
                "</mc:Choice><mc:Fallback>" + Reference(fallbackId) + "</mc:Fallback></mc:AlternateContent>";
        }, RegexOptions.None, TimeSpan.FromSeconds(5)));

        using var document = WordDocument.Load(path);
        WordChart chart = Assert.Single(document.AllCharts);
        Assert.Equal("Sales", chart.Title);
        Assert.True(chart.TrySetSeriesName(0, "Edited"));
    }

    [Fact]
    public void SeriesInTheSelectedCompatibilityBranchAreEditedAndCounted() {
        string path = CreateBarChart("alternate-series.docx");
        RewriteChartPart(path, xml => Regex.Replace(xml, "(<c:ser>.*?</c:ser>)", "<mc:AlternateContent xmlns:mc=\"" + McNs + "\" xmlns:v=\"urn:example:vendor\">" +
            "<mc:Choice Requires=\"v\"><c:ser><c:idx val=\"0\" /><c:order val=\"0\" /><c:tx><c:v>Choice</c:v></c:tx></c:ser></mc:Choice>" +
            "<mc:Fallback>$1</mc:Fallback></mc:AlternateContent>", RegexOptions.Singleline, TimeSpan.FromSeconds(5)));

        using (var document = WordDocument.Load(path)) {
            WordChart chart = Assert.Single(document.AllCharts);
            Assert.True(chart.TrySetSeriesName(0, "Edited"));
            chart.AddBar("Added", new[] { 3, 4 }, OfficeColor.Black);
            document.Save();
        }

        string xml = Assert.Single(ReadChartParts(path));
        string choice = Regex.Match(xml, "<mc:Choice.*?</mc:Choice>", RegexOptions.Singleline, TimeSpan.FromSeconds(5)).Value;
        string fallback = Regex.Match(xml, "<mc:Fallback>.*?</mc:Fallback>", RegexOptions.Singleline, TimeSpan.FromSeconds(5)).Value;
        Assert.Contains("<c:v>Choice</c:v>", choice, StringComparison.Ordinal);
        Assert.Contains("<c:v>Edited</c:v>", fallback, StringComparison.Ordinal);
        Match added = Regex.Match(xml, "<c:ser><c:idx val=\"(\\d+)\"(?:(?!</c:ser>).)*<c:v>Added</c:v>", RegexOptions.Singleline, TimeSpan.FromSeconds(5));
        Assert.True(added.Success);
        Assert.NotEqual("0", added.Groups[1].Value);
    }

    [Fact]
    public void MissingPieLabelsAreSavedAsBlankCategories() {
        string path = Path.Combine(_directory, "blank-label.docx");
        using (var document = WordDocument.Create(path)) {
            document.AddChart("Share").AddPie(null!, 1).AddPie("B", 2);
            document.Save();
        }

        string categories = Regex.Match(Assert.Single(ReadChartParts(path)), "<c:cat>.*?</c:cat>", RegexOptions.Singleline, TimeSpan.FromSeconds(5)).Value;
        Assert.Matches("<c:pt idx=\"0\"><c:v\\s*(/>|></c:v>)", categories);
        Assert.Contains("<c:v>B</c:v>", categories, StringComparison.Ordinal);
    }

    [Fact]
    public void SeriesInsideVendorMarkupDoNotTakePointStyles() {
        string path = CreateBarChart("vendor-series.docx");
        RewriteChartPart(path, xml => ReplaceFirst(xml, "<c:ser>",
            "<v:meta xmlns:v=\"urn:example:vendor\"><c:ser><c:idx val=\"0\" /><c:order val=\"0\" /><c:tx><c:v>Decoy</c:v></c:tx></c:ser></v:meta><c:ser>"));

        using (var document = WordDocument.Load(path)) {
            Assert.Single(document.AllCharts).SetDataPointStyle(0, 1, new OfficeChartPointStyle(fillColor: OfficeColor.Parse("#FF0000")));
            document.Save();
        }

        string xml = Assert.Single(ReadChartParts(path));
        string vendor = Regex.Match(xml, "<v:meta.*?</v:meta>", RegexOptions.Singleline, TimeSpan.FromSeconds(5)).Value;
        Assert.Contains("<c:v>Decoy</c:v>", vendor, StringComparison.Ordinal);
        Assert.DoesNotContain("<c:dPt>", vendor, StringComparison.Ordinal);
        Assert.Contains("<c:dPt>", xml.Replace(vendor, string.Empty), StringComparison.Ordinal);
    }

    [Fact]
    public void ValuesInsideVendorMarkupDoNotScaleTheCountAxis() {
        string path = CreateBarChart("vendor-count.docx");
        RewriteChartPart(path, xml => ReplaceFirst(xml, "<c:ser>",
            "<v:meta xmlns:v=\"urn:example:vendor\"><c:val><c:numLit><c:ptCount val=\"1\" /><c:pt idx=\"0\"><c:v>999999999</c:v></c:pt></c:numLit></c:val></v:meta><c:ser>"));

        using (var document = WordDocument.Load(path)) {
            Assert.Single(document.AllCharts).SetCountAxis();
            document.Save();
        }

        string scaling = Regex.Match(Assert.Single(ReadChartParts(path)), "<c:valAx\\b.*?</c:scaling>", RegexOptions.Singleline, TimeSpan.FromSeconds(5)).Value;
        Assert.Contains("<c:max val=\"3\"", scaling, StringComparison.Ordinal);
    }

    [Fact]
    public void SeriesTitleValuesInsideVendorMarkupDoNotPickTheOutcomeColor() {
        string path = Path.Combine(_directory, "vendor-series-title.docx");
        using (var document = WordDocument.Create(path)) {
            WordChart chart = document.AddChart("Results");
            chart.AddCategories(new[] { "Mon", "Tue" });
            chart.AddBar("Passed", new[] { 1, 2 }, OfficeColor.Black);
            document.Save();
        }
        RewriteChartPart(path, xml => xml.Replace("<c:tx><c:v>Passed</c:v></c:tx>",
            "<c:tx><v:meta xmlns:v=\"urn:example:vendor\"><c:v>Failed</c:v></v:meta><c:v>Passed</c:v></c:tx>"));

        using (var document = WordDocument.Load(path)) {
            WordChart chart = Assert.Single(document.AllCharts);
            chart.ApplyPalette(WordChart.WordChartPalette.Professional);
            document.Save();
        }

        string series = Regex.Match(Assert.Single(ReadChartParts(path)), "<c:ser>.*?</c:ser>", RegexOptions.Singleline, TimeSpan.FromSeconds(5)).Value;
        Assert.Contains("2FB344", series, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("F76707", series, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PieSliceColorGoesBeforeTheSliceOutline() {
        string path = Path.Combine(_directory, "slice-outline.docx");
        using (var document = WordDocument.Create(path)) {
            WordChart chart = document.AddChart("Share");
            chart.AddPie("A", 1).AddPie("B", 2);
            chart.SetPieSliceColor(0, OfficeColor.Parse("#FF0000"));
            document.Save();
        }
        RewriteChartPart(path, xml => Regex.Replace(xml, "(<c:dPt>.*?</a:solidFill>)", "$1<a:ln w=\"9525\" />",
            RegexOptions.Singleline, TimeSpan.FromSeconds(5)));

        using (var document = WordDocument.Load(path)) {
            Assert.Single(document.AllCharts).SetPieSliceColor(0, OfficeColor.Parse("#0000FF"));
            document.Save();
        }

        string point = Regex.Match(Assert.Single(ReadChartParts(path)), "<c:dPt>.*?</c:dPt>", RegexOptions.Singleline, TimeSpan.FromSeconds(5)).Value;
        Assert.Contains("0000FF", point, StringComparison.Ordinal);
        Assert.True(point.IndexOf("<a:solidFill>", StringComparison.Ordinal) < point.IndexOf("<a:ln ", StringComparison.Ordinal));
    }

    private static string ReplaceFirst(string text, string value, string replacement) {
        int index = text.IndexOf(value, StringComparison.Ordinal);
        Assert.True(index >= 0);
        return text.Substring(0, index) + replacement + text.Substring(index + value.Length);
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
