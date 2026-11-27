using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

/// <summary>
/// Chart reads look for chart content outside extension markup, as the SDK did when it read extensions as unknown
/// elements, refuse an oversized chart style part whichever host holds it, and chart edits that hit markup they can't
/// read change nothing.
/// </summary>
public sealed class PowerPointChartXmlLookupTests {
    private const string PlotArea =
        "<c:plotArea xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\""
        + " xmlns:c15=\"http://schemas.microsoft.com/office/drawing/2012/chart\""
        + " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">"
        + "<c:barChart><c:barDir val=\"col\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:cat><c:strRef><c:f>Sheet1!$A$2</c:f></c:strRef></c:cat></c:ser>"
        + "<c:extLst><c:ext uri=\"{02D57815-91ED-43cb-92C2-25804820EDAC}\"><c15:filteredBarSeries><c15:ser><c:idx val=\"1\"/><c:order val=\"1\"/>"
        + "<c:cat><c:numRef><c:f>Sheet1!$B$2</c:f></c:numRef></c:cat></c15:ser></c15:filteredBarSeries></c:ext></c:extLst></c:barChart>"
        + "<mc:AlternateContent><mc:Choice Requires=\"c15\"><c:cat/></mc:Choice><mc:Fallback><c:cat/></mc:Fallback></mc:AlternateContent>"
        + "</c:plotArea>";

    [Fact]
    public void ChartLookupsLeaveOutExtensionMarkupOnTheTreeAndOnSdkElements() {
        byte[] bytes = Encoding.UTF8.GetBytes(PlotArea);
        OfficeXmlTree tree = OfficeXmlTree.Parse(bytes, bytes.Length, new OfficeXmlNamespaces.Table());
        foreach (OfficeXmlNode plot in new[] { new OfficeXmlNode(tree, tree.Root), OfficeXmlNode.FromSdk(new C.PlotArea(PlotArea)) }) {
            OfficeXmlNode category = Assert.Single(plot.ContentDescendants(OfficeXmlNamespaces.Chart, "cat"));
            Assert.Equal("Sheet1!$A$2", category.InnerText);
            Assert.Equal(category, plot.FirstDescendant(OfficeXmlNamespaces.Chart, "cat"));
            Assert.Single(plot.DescendantsWhere(element => element.Is(OfficeXmlNamespaces.Chart, "f")));
            Assert.DoesNotContain(plot.AllDescendants(), element => element.LocalName is "extLst" or "AlternateContent");

            // A read that means to look inside extensions still can.
            Assert.Equal(4, plot.DescendantsWhere(element => element.Is(OfficeXmlNamespaces.Chart, "cat"), includeExtensionMarkup: true).Count);
            Assert.Single(plot.DescendantsWhere(element => element.Is(OfficeXmlNamespaces.Chart, "cat"), includeExtensionMarkup: true, limit: 1));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OversizedChartStyleIsRefused(bool sdk) {
        PowerPointChart control = Chart(LabelledDeck(), sdk, out PowerPointDocument controlDeck);
        using (controlDeck) Assert.True(control.TryGetOfficeSnapshot(out _));

        byte[] oversized = LabelledDeck();
        using (var stream = new MemoryStream()) {
            stream.Write(oversized, 0, oversized.Length);
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
                ZipArchiveEntry entry = archive.Entries.Single(item => item.FullName.StartsWith("ppt/charts/style", StringComparison.Ordinal));
                XDocument style;
                using (Stream read = entry.Open()) style = XDocument.Load(read);
                XNamespace pad = "urn:officeimo:test";
                style.Root!.Add(new XAttribute(XNamespace.Xmlns + "pad", pad.NamespaceName), new XAttribute(pad + "fill", new string('x', 140000)));
                entry.Delete();
                using Stream write = archive.CreateEntry(entry.FullName).Open();
                style.Save(write);
            }
            oversized = stream.ToArray();
        }

        PowerPointChart chart = Chart(oversized, sdk, out PowerPointDocument presentation);
        using (presentation) Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void CachePointsInsideExtensionMarkupAreNotData() {
        byte[] deck = RewriteChart(LabelledDeck(), xml => ReplaceFirst(xml, "</c:strCache>",
            "<c:extLst><c:ext uri=\"{00000000-0000-0000-0000-000000000001}\"><c:pt idx=\"5\"><c:v>Z</c:v></c:pt></c:ext></c:extLst></c:strCache>"));
        PowerPointChart chart = Chart(deck, sdk: false, out PowerPointDocument presentation);
        using (presentation) {
            Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
            Assert.Equal(new[] { "A", "B" }, snapshot.Data.Categories);
        }
    }

    [Theory]
    [InlineData("<a:scrgbClr r=\"100%\" g=\"0%\" b=\"0%\"/>")]
    [InlineData("<a:scrgbClr r=\"100000\" g=\"0\" b=\"0\"/>")]
    [InlineData("<a:hslClr hue=\"0\" sat=\"100%\" lum=\"50%\"/>")]
    public void PercentageColorComponentsResolveInEitherForm(string color) {
        string fill = "<a:solidFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" + color + "</a:solidFill>";
        byte[] bytes = Encoding.UTF8.GetBytes(fill);
        OfficeXmlTree tree = OfficeXmlTree.Parse(bytes, bytes.Length, new OfficeXmlNamespaces.Table());

        Assert.Equal(OfficeColor.FromRgb(0xFF, 0x00, 0x00), OfficeOpenXmlThemeColorResolver.ResolveColor(new OfficeXmlNode(tree, tree.Root), default));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AnUnreadableSeriesIndexLeavesEverySeriesUnchanged(bool sdk) {
        byte[] deck = RewriteChart(TwoSeriesDeck(), xml => ReplaceFirst(xml, "<c:idx val=\"0\"/>", "<c:idx val=\"bad\"/>"));
        PowerPointChart chart = Chart(deck, sdk, out PowerPointDocument presentation);
        using (presentation) {
            Assert.Throws<FormatException>(() => chart.SetSeriesFillColor(0, OfficeColor.FromRgb(0x12, 0x34, 0x56)));
            Assert.DoesNotContain("123456", ChartXml(presentation.ToBytes()), StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AnUnreadableKeptAxisBoundRejectsAPartialScaleEdit(bool sdk) {
        byte[] scaled;
        using (var source = PowerPointDocument.Load(new MemoryStream(LabelledDeck()))) {
            source.Slides[0].Shapes.OfType<PowerPointChart>().Single().SetValueAxisScale(minimum: 0, maximum: 10);
            scaled = source.ToBytes();
        }
        byte[] deck = RewriteChart(scaled, xml => ReplaceFirst(xml, "<c:max val=\"10\"/>", "<c:max val=\"bad\"/>"));
        PowerPointChart chart = Chart(deck, sdk, out PowerPointDocument presentation);
        using (presentation) Assert.Throws<FormatException>(() => chart.SetValueAxisScale(minimum: 1));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AFailedTrendlineEditKeepsTheCurrentTrendline(bool sdk) {
        byte[] trended;
        using (var source = PowerPointDocument.Load(new MemoryStream(LabelledDeck()))) {
            source.Slides[0].Shapes.OfType<PowerPointChart>().Single().SetSeriesTrendline(0, OfficeChartTrendlineType.Linear);
            trended = source.ToBytes();
        }
        PowerPointChart chart = Chart(trended, sdk, out PowerPointDocument presentation);
        using (presentation) {
            Assert.Throws<OverflowException>(() => chart.SetSeriesTrendline(0, OfficeChartTrendlineType.Exponential,
                lineColor: OfficeColor.FromRgb(0x12, 0x34, 0x56), lineWidth: OfficeLength.Points(200_000)));
            string xml = ChartXml(presentation.ToBytes());
            Assert.Contains("trendlineType val=\"linear\"", xml, StringComparison.Ordinal);
            Assert.DoesNotContain("exp", xml, StringComparison.Ordinal);
        }
    }

    private static string ReplaceFirst(string text, string value, string replacement) {
        int index = text.IndexOf(value, StringComparison.Ordinal);
        Assert.True(index >= 0, value);
        return text.Substring(0, index) + replacement + text.Substring(index + value.Length);
    }

    private static string ChartXml(byte[] deck) {
        using var archive = new ZipArchive(new MemoryStream(deck));
        using var reader = new StreamReader(archive.Entries.Single(item => item.FullName.StartsWith("ppt/charts/chart", StringComparison.Ordinal)).Open());
        return reader.ReadToEnd();
    }

    private static byte[] RewriteChart(byte[] deck, Func<string, string> change) {
        using var stream = new MemoryStream();
        stream.Write(deck, 0, deck.Length);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry entry = archive.Entries.Single(item => item.FullName.StartsWith("ppt/charts/chart", StringComparison.Ordinal));
            string xml;
            using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
            string name = entry.FullName;
            entry.Delete();
            using var writer = new StreamWriter(archive.CreateEntry(name).Open());
            // The SDK writes "<x />" and the package writer "<x/>"; the edits match the second.
            writer.Write(change(xml.Replace(" />", "/>")));
        }
        return stream.ToArray();
    }

    private static byte[] TwoSeriesDeck() {
        using var stream = new MemoryStream();
        using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
            presentation.AddSlide().AddChart(OfficeChartType.Line, new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeSeries("First", new[] { 1d, 2d }), new OfficeSeries("Second", new[] { 3d, 4d }) }));
            presentation.Save();
        }
        return stream.ToArray();
    }

    private static PowerPointChart Chart(byte[] deck, bool sdk, out PowerPointDocument presentation) {
        presentation = PowerPointDocument.Load(new MemoryStream(deck, writable: false));
        if (sdk) _ = presentation.OpenXmlDocument;
        return presentation.Slides.Single().Shapes.OfType<PowerPointChart>().Single();
    }

    private static byte[] LabelledDeck() {
        using var stream = new MemoryStream();
        using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
            presentation.AddSlide().AddChart(OfficeChartType.ColumnClustered,
                new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeSeries("Values", new[] { 3d, 4d }) })).SetDataLabels(showValue: true);
            presentation.Save();
        }
        return stream.ToArray();
    }
}
