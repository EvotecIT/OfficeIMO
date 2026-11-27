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
/// elements, and refuse an oversized chart style part whichever host holds it.
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
