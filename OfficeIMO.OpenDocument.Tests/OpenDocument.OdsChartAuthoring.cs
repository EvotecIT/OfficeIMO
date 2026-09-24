using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdsChartAuthoringTests {
    [Theory]
    [InlineData(OdsChartType.Column, "chart:bar", false)]
    [InlineData(OdsChartType.Bar, "chart:bar", true)]
    [InlineData(OdsChartType.Line, "chart:line", null)]
    public void AuthoredChartReopensWithSourceRanges(OdsChartType type, string chartClass, bool? vertical) {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Month");
        sheet.Cell(0, 1).SetString("Sales");
        sheet.Cell(1, 0).SetString("Jan");
        sheet.Cell(1, 1).SetNumber(10);
        sheet.Cell(2, 0).SetString("Feb");
        sheet.Cell(2, 1).SetNumber(20);
        OdsChart chart = sheet.AddChart(type, "Data.$A$2:.$A$3",
            new[] { new OdsChartSeries("Data.$B$2:.$B$3", "Data.$B$1") },
            4, 2, OdfRect.FromCentimeters(1, 1, 10, 6), "Sales");

        Assert.Equal(chartClass, chart.ChartClass);
        Assert.Equal(vertical, chart.VerticalBars);
        OdsDocument reopened = OdsDocument.Load(new MemoryStream(document.ToBytes()));
        OdsChart actual = Assert.Single(reopened.GetSheet("Data")!.Charts);
        Assert.Equal(chartClass, actual.ChartClass);
        Assert.Equal("Sales", actual.Title);
        Assert.Equal("Data.$A$2:.$A$3", actual.CategoriesAddress);
        Assert.Equal("Data.$B$2:.$B$3", Assert.Single(actual.Series).ValuesAddress);
        Assert.Equal(4, actual.AnchorRow);
        Assert.Equal(2, actual.AnchorColumn);
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void ChartAuthoringRejectsNonlocalAndMismatchedRangesBeforeWriting() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        Assert.Throws<ArgumentException>(() => sheet.AddChart(OdsChartType.Line, "Missing.$A$1:.$A$2",
            new[] { new OdsChartSeries("Data.$B$1:.$B$2") }, 0, 0,
            OdfRect.FromCentimeters(1, 1, 10, 6)));
        Assert.Throws<ArgumentException>(() => sheet.AddChart(OdsChartType.Line, "Data.$A$1:.$A$2",
            new[] { new OdsChartSeries("Data.$B$1:.$B$3") }, 0, 0,
            OdfRect.FromCentimeters(1, 1, 10, 6)));
        Assert.Empty(sheet.Charts);
    }

    [Fact]
    public void VersionRewriteKeepsManifestOnlyEmbeddedChartDirectory() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "microsoft-excel-column-chart.ods");
        OdsDocument imported = OdsDocument.Load(path);
        byte[] rewritten = imported.ToBytes(new OdfSaveOptions {
            CompatibilityProfile = OdfCompatibilityProfile.Odf13
        });
        OdsDocument reopened = OdsDocument.Load(new MemoryStream(rewritten));
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace manifest = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
        XDocument embedded = XDocument.Parse(Encoding.UTF8.GetString(
            reopened.GetPackageEntryBytes("Object 1/content.xml")));
        XDocument listing = XDocument.Parse(Encoding.UTF8.GetString(
            reopened.GetPackageEntryBytes("META-INF/manifest.xml")));
        Assert.Equal("1.3", (string?)embedded.Root!.Attribute(office + "version"));
        XElement directory = Assert.Single(listing.Descendants(manifest + "file-entry"),
            entry => (string?)entry.Attribute(manifest + "full-path") == "Object 1/");
        Assert.Equal("application/vnd.oasis.opendocument.chart",
            (string?)directory.Attribute(manifest + "media-type"));
        Assert.Single(reopened.GetSheet("Data")!.Charts);
    }
}
