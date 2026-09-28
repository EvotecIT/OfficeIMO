using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.OpenDocument;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdsChartAuthoringTests {
    [Theory]
    [InlineData(OdsChartType.Pie)]
    [InlineData(OdsChartType.Doughnut)]
    public void RadialPointStylesSurviveNativePackageReopen(OdsChartType type) {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        for (int index = 0; index < 3; index++) {
            sheet.Cell(index, 0).SetString("Category " + index);
            sheet.Cell(index, 1).SetNumber(index + 1);
        }
        OfficeChartPointStyle?[] styles = {
            new(fillColor: OfficeColor.Parse("#228844")),
            new(fillColor: OfficeColor.Parse("#FFF0DD"), hatch: OfficeChartHatchPattern.WideForwardDiagonal,
                hatchColor: OfficeColor.Parse("#D97706")),
            new(noFill: true, outlineColor: OfficeColor.Parse("#445566"), outlineWidth: 2,
                showOutline: true, outlineJoin: OfficeStrokeLineJoin.Round)
        };
        sheet.AddChart(type, "Data.$A$1:.$A$3",
            new[] { new OdsChartSeries("Data.$B$1:.$B$3").WithPointStyles(styles) },
            4, 2, OdfRect.FromCentimeters(1, 1, 10, 6), "Status");

        XDocument chartPart = XDocument.Parse(Encoding.UTF8.GetString(
            document.GetPackageEntryBytes("Object 1/content.xml")));
        XElement seriesPart = Assert.Single(chartPart.Descendants(OdfNamespaces.Chart + "series"));
        Assert.Equal("chart:circle", (string?)seriesPart.Attribute(OdfNamespaces.Chart + "class"));
        string seriesStyle = Assert.IsType<string>((string?)seriesPart.Attribute(OdfNamespaces.Chart + "style-name"));
        Assert.Contains(chartPart.Descendants(OdfNamespaces.Style + "style"), definition =>
            (string?)definition.Attribute(OdfNamespaces.Style + "name") == seriesStyle &&
            (string?)definition.Attribute(OdfNamespaces.Style + "family") == "chart");

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(document.ToBytes()));
        OdsChart chart = Assert.Single(reopened.GetSheet("Data")!.Charts);
        Assert.Equal(type == OdsChartType.Pie ? "chart:circle" : "chart:ring", chart.ChartClass);
        OfficeChartPointStyle?[] actual = Assert.Single(chart.Series).PointStyles!.ToArray();
        Assert.Equal(styles.Select(style => style!.FillColor), actual.Select(style => style!.FillColor));
        Assert.Equal(OfficeChartHatchPattern.WideForwardDiagonal, actual[1]!.Hatch);
        Assert.Equal(OfficeColor.Parse("#D97706"), actual[1]!.HatchColor);
        Assert.True(actual[2]!.NoFill);
        Assert.Equal(OfficeColor.Parse("#445566"), actual[2]!.OutlineColor);
        Assert.True(actual[2]!.ShowOutline);
        Assert.Equal(OfficeStrokeLineJoin.Round, actual[2]!.OutlineJoin);
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void LibreOfficePiePointStylesProjectFromIndependentChartParts() {
        string fixture = Path.Combine(AppContext.BaseDirectory, "ChartProducerReferences", "status-pie.odt");
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        for (int index = 0; index < 3; index++) {
            sheet.Cell(index, 0).SetString("Category " + index);
            sheet.Cell(index, 1).SetNumber(index + 1);
        }
        sheet.AddChart(OdsChartType.Pie, "Data.$A$1:.$A$3",
            new[] { new OdsChartSeries("Data.$B$1:.$B$3") },
            4, 2, OdfRect.FromCentimeters(1, 1, 10, 6));
        using (var stream = File.OpenRead(fixture))
        using (var package = new ZipArchive(stream, ZipArchiveMode.Read)) {
            foreach (string file in new[] { "content.xml", "styles.xml" }) {
                using var source = package.GetEntry("Object 1/" + file)!.Open();
                using var bytes = new MemoryStream();
                source.CopyTo(bytes);
                document.Package.AddOrReplaceEntry("Object 1/" + file, bytes.ToArray(), "text/xml");
            }
        }
        OdsChart chart = Assert.Single(sheet.Charts);
        Assert.Equal("chart:circle", chart.ChartClass);
        OfficeChartPointStyle?[] styles = Assert.Single(chart.Series).PointStyles!.ToArray();
        Assert.Equal(3, styles.Length);
        Assert.Equal(OfficeColor.Parse("#228844"), styles[0]!.FillColor);
        Assert.Equal(OfficeChartHatchPattern.ForwardDiagonal, styles[1]!.Hatch);
        Assert.Equal(OfficeColor.Parse("#FFF0DD"), styles[1]!.FillColor);
        Assert.Equal(OfficeColor.Parse("#D97706"), styles[1]!.HatchColor);
        Assert.True(styles[2]!.NoFill);
        Assert.Equal(OfficeColor.Parse("#445566"), styles[2]!.OutlineColor);
        Assert.Equal(0.07 * 72 / 2.54, styles[2]!.OutlineWidth!.Value, 6);
    }

    [Fact]
    public void ImportedPointFillInheritsSeriesFillMode() {
        OdsChart chart = ReadProducerChart(content => {
            XElement series = Assert.Single(content.Descendants(OdfNamespaces.Chart + "series"));
            string name = Assert.IsType<string>((string?)series.Attribute(OdfNamespaces.Chart + "style-name"));
            XElement definition = Assert.Single(content.Descendants(OdfNamespaces.Style + "style"),
                item => (string?)item.Attribute(OdfNamespaces.Style + "name") == name);
            definition.Element(OdfNamespaces.Style + "graphic-properties")!
                .SetAttributeValue(OdfNamespaces.Draw + "fill", "none");
        });

        OfficeChartPointStyle?[] styles = Assert.Single(chart.Series).PointStyles!.ToArray();
        Assert.True(styles[0]!.NoFill);
        Assert.Null(styles[0]!.FillColor);
        Assert.Equal(OfficeChartHatchPattern.ForwardDiagonal, styles[1]!.Hatch);
    }

    [Fact]
    public void ImportedPointAppearanceDoesNotInheritPlotBackgroundGraphics() {
        OdsChart chart = ReadProducerChart(content => {
            XElement plot = Assert.Single(content.Descendants(OdfNamespaces.Chart + "plot-area"));
            var style = new XElement(OdfNamespaces.Style + "style",
                new XAttribute(OdfNamespaces.Style + "name", "PlotBackgroundProbe"),
                new XAttribute(OdfNamespaces.Style + "family", "chart"),
                new XElement(OdfNamespaces.Style + "graphic-properties",
                    new XAttribute(OdfNamespaces.Draw + "fill", "solid"),
                    new XAttribute(OdfNamespaces.Draw + "fill-color", "#EE11CC"),
                    new XAttribute(OdfNamespaces.Draw + "stroke", "solid"),
                    new XAttribute(OdfNamespaces.Svg + "stroke-color", "#FF0000")));
            content.Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Add(style);
            plot.SetAttributeValue(OdfNamespaces.Chart + "style-name", "PlotBackgroundProbe");
        });

        OfficeChartPointStyle?[] points = Assert.Single(chart.Series).PointStyles!.ToArray();
        Assert.Equal(OfficeColor.Parse("#228844"), points[0]!.FillColor);
        Assert.All(points, point => Assert.NotEqual(OfficeColor.Parse("#FF0000"), point?.OutlineColor));
    }

    [Fact]
    public void ImportedRepeatedPointStylesHaveAnAggregateSeriesBound() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("A");
        sheet.Cell(0, 1).SetNumber(1);
        sheet.AddChart(OdsChartType.Pie, "Data.$A$1",
            new[] { new OdsChartSeries("Data.$B$1").WithPointStyles(
                new OfficeChartPointStyle?[] { new(fillColor: OfficeColor.Parse("#228844")) }) },
            4, 2, OdfRect.FromCentimeters(1, 1, 10, 6));
        Assert.Single(sheet.Charts);
        const string partPath = "Object 1/content.xml";
        XDocument xml = document.Package.GetXml(partPath);
        XElement plot = Assert.Single(xml.Descendants(OdfNamespaces.Chart + "plot-area"));
        XElement series = Assert.Single(plot.Elements(OdfNamespaces.Chart + "series"));
        series.SetAttributeValue(OdfNamespaces.Chart + "values-cell-range-address", "Data.$B$1:.$B$4096");
        Assert.Single(series.Elements(OdfNamespaces.Chart + "data-point"))
            .SetAttributeValue(OdfNamespaces.Chart + "repeated", "4096");
        for (int index = 1; index <= 256; index++) plot.Add(new XElement(series));
        document.Package.AddOrReplaceEntry(partPath,
            Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)), "text/xml");

        Assert.Empty(sheet.Charts);
    }

    [Fact]
    public void UnsupportedProducerHatchSpacingRemainsUnprojected() {
        OdsChart chart = ReadProducerChart(styles => {
            Assert.Single(styles.Descendants(OdfNamespaces.Draw + "hatch"))
                .SetAttributeValue(OdfNamespaces.Draw + "distance", "0.8cm");
        }, editStyles: true);

        Assert.Null(Assert.Single(chart.Series).PointStyles);
    }

    [Fact]
    public void ImportedPointStylesIncludeImplicitTrailingPoints() {
        OdsChart chart = ReadProducerChart(content => {
            XElement series = Assert.Single(content.Descendants(OdfNamespaces.Chart + "series"));
            foreach (XElement point in series.Elements(OdfNamespaces.Chart + "data-point").Skip(1).ToArray())
                point.Remove();
        });

        OfficeChartPointStyle?[] points = Assert.Single(chart.Series).PointStyles!.ToArray();
        Assert.Equal(3, points.Length);
        Assert.Equal(OfficeColor.Parse("#228844"), points[0]!.FillColor);
        Assert.Null(points[1]);
        Assert.Null(points[2]);
    }

    [Theory]
    [InlineData("shadow", "visible")]
    [InlineData("opacity", "50%")]
    public void UnsupportedPointEffectsRemainUnprojected(string attribute, string value) {
        OdsChart chart = ReadProducerChart(content => {
            XElement series = Assert.Single(content.Descendants(OdfNamespaces.Chart + "series"));
            XElement point = series.Element(OdfNamespaces.Chart + "data-point")!;
            string name = (string)point.Attribute(OdfNamespaces.Chart + "style-name")!;
            XElement definition = Assert.Single(content.Descendants(OdfNamespaces.Style + "style"),
                item => (string?)item.Attribute(OdfNamespaces.Style + "name") == name);
            definition.Element(OdfNamespaces.Style + "graphic-properties")!
                .SetAttributeValue(OdfNamespaces.Draw + attribute, value);
        });

        Assert.Null(Assert.Single(chart.Series).PointStyles);
    }

    [Fact]
    public void ExplodedProducerSliceWithSupportedFillRemainsUnprojected() {
        OdsChart chart = ReadProducerChart(content => {
            XElement series = Assert.Single(content.Descendants(OdfNamespaces.Chart + "series"));
            string name = (string)series.Element(OdfNamespaces.Chart + "data-point")!
                .Attribute(OdfNamespaces.Chart + "style-name")!;
            XElement definition = Assert.Single(content.Descendants(OdfNamespaces.Style + "style"),
                item => (string?)item.Attribute(OdfNamespaces.Style + "name") == name);
            definition.Element(OdfNamespaces.Style + "chart-properties")!
                .SetAttributeValue(OdfNamespaces.Chart + "pie-offset", "25");
        });

        Assert.Null(Assert.Single(chart.Series).PointStyles);
    }

    [Fact]
    public void ExplicitZeroPieOffsetOverridesInheritedDefault() {
        var defaultStyle = new XElement(OdfNamespaces.Style + "default-style",
            new XElement(OdfNamespaces.Style + "chart-properties",
                new XAttribute(OdfNamespaces.Chart + "pie-offset", "25")));
        var seriesStyle = new XElement(OdfNamespaces.Style + "style",
            new XElement(OdfNamespaces.Style + "chart-properties",
                new XAttribute(OdfNamespaces.Chart + "pie-offset", "0")));
        var series = new XElement(OdfNamespaces.Chart + "series",
            new XAttribute(OdfNamespaces.Chart + "style-name", "SeriesStyle"));
        Assert.False(OdsChartPointStyles.HasUnprojectedSeriesPieOffset(series,
            name => name == "SeriesStyle" ? seriesStyle : null, defaultStyle));
    }

    [Fact]
    public void ImportedSolidHatchAcceptsXmlBooleanOne() {
        OdsChart chart = ReadProducerChart(content => {
            XElement series = Assert.Single(content.Descendants(OdfNamespaces.Chart + "series"));
            XElement point = series.Elements(OdfNamespaces.Chart + "data-point").Skip(1).First();
            string name = (string)point.Attribute(OdfNamespaces.Chart + "style-name")!;
            XElement definition = Assert.Single(content.Descendants(OdfNamespaces.Style + "style"),
                item => (string?)item.Attribute(OdfNamespaces.Style + "name") == name);
            definition.Element(OdfNamespaces.Style + "graphic-properties")!
                .SetAttributeValue(OdfNamespaces.Draw + "fill-hatch-solid", "1");
        });

        Assert.Equal(OfficeChartHatchPattern.ForwardDiagonal,
            Assert.Single(chart.Series).PointStyles![1]!.Hatch);
    }

    [Fact]
    public void SubMillipointOutlineIsRejectedBeforeWritingChart() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("A");
        sheet.Cell(0, 1).SetNumber(1);
        Assert.Throws<NotSupportedException>(() => sheet.AddChart(OdsChartType.Pie, "Data.$A$1",
            new[] { new OdsChartSeries("Data.$B$1").WithPointStyles(
                new OfficeChartPointStyle?[] {
                    new(outlineColor: OfficeColor.Parse("#445566"), outlineWidth: 0.0004,
                        showOutline: true)
                }) }, 0, 2, OdfRect.FromCentimeters(0, 0, 10, 6)));
        Assert.Empty(sheet.Charts);
    }

    [Fact]
    public void StyledLineChartRequiresVisiblePointSymbols() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("A");
        sheet.Cell(0, 1).SetNumber(1);
        Assert.Throws<NotSupportedException>(() => sheet.AddChart(OdsChartType.Line, "Data.$A$1",
            new[] { new OdsChartSeries("Data.$B$1").WithPointStyles(
                new OfficeChartPointStyle?[] { new(fillColor: OfficeColor.Parse("#228844")) }) },
            0, 2, OdfRect.FromCentimeters(0, 0, 10, 6)));
        Assert.Empty(sheet.Charts);
    }

    [Fact]
    public void OutlineOnlyPointDoesNotInheritChartBackgroundFill() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        for (int index = 0; index < 2; index++) {
            sheet.Cell(index, 0).SetString("Category " + index);
            sheet.Cell(index, 1).SetNumber(index + 1);
        }
        sheet.AddChart(OdsChartType.Pie, "Data.$A$1:.$A$2",
            new[] { new OdsChartSeries("Data.$B$1:.$B$2").WithPointStyles(
                new OfficeChartPointStyle?[] {
                    new(outlineColor: OfficeColor.Parse("#445566"), showOutline: true), null
                }) }, 0, 2, OdfRect.FromCentimeters(0, 0, 10, 6));

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(document.ToBytes()));
        OfficeChartPointStyle?[] points = Assert.Single(
            Assert.Single(reopened.GetSheet("Data")!.Charts).Series).PointStyles!.ToArray();
        Assert.False(points[0]!.NoFill);
        Assert.Null(points[0]!.FillColor);
        Assert.True(points[0]!.ShowOutline);
        Assert.Equal(OfficeColor.Parse("#445566"), points[0]!.OutlineColor);
        Assert.Null(points[1]);
    }

    private static OdsChart ReadProducerChart(Action<XDocument> edit,
        bool editStyles = false) {
        string fixture = Path.Combine(AppContext.BaseDirectory, "ChartProducerReferences", "status-pie.odt");
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        for (int index = 0; index < 3; index++) {
            sheet.Cell(index, 0).SetString("Category " + index);
            sheet.Cell(index, 1).SetNumber(index + 1);
        }
        sheet.AddChart(OdsChartType.Pie, "Data.$A$1:.$A$3",
            new[] { new OdsChartSeries("Data.$B$1:.$B$3") },
            4, 2, OdfRect.FromCentimeters(1, 1, 10, 6));
        using (var stream = File.OpenRead(fixture))
        using (var package = new ZipArchive(stream, ZipArchiveMode.Read)) {
            foreach (string file in new[] { "content.xml", "styles.xml" }) {
                using var source = package.GetEntry("Object 1/" + file)!.Open();
                XDocument xml = XDocument.Load(source);
                if ((file == "styles.xml") == editStyles) edit(xml);
                document.Package.AddOrReplaceEntry("Object 1/" + file,
                    Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)), "text/xml");
            }
        }
        return Assert.Single(sheet.Charts);
    }

    [Theory]
    [InlineData(OdsChartType.Column, "chart:bar", false)]
    [InlineData(OdsChartType.Bar, "chart:bar", true)]
    [InlineData(OdsChartType.Line, "chart:line", null)]
    [InlineData(OdsChartType.Pie, "chart:circle", null)]
    [InlineData(OdsChartType.Doughnut, "chart:ring", null)]
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

        XDocument chartPart = XDocument.Parse(Encoding.UTF8.GetString(
            document.GetPackageEntryBytes("Object 1/content.xml")));
        XElement background = Assert.Single(chartPart.Descendants(OdfNamespaces.Style + "style"),
            definition => (string?)definition.Attribute(OdfNamespaces.Style + "name") == "ChartStyle")
            .Element(OdfNamespaces.Style + "graphic-properties")!;
        Assert.Equal("solid", (string?)background.Attribute(OdfNamespaces.Draw + "fill"));
        Assert.Equal("#FFFFFF", (string?)background.Attribute(OdfNamespaces.Draw + "fill-color"));

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
    public void ChartAuthoringRejectsUnqualifiedCrossSheetRangeEnd() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet chartSheet = document.AddSheet("Summary");
        document.AddSheet("Data");
        Assert.Throws<ArgumentException>(() => chartSheet.AddChart(OdsChartType.Line,
            "Data.$A$1:.$A$2", new[] { new OdsChartSeries("Data.$B$1:Data.$B$2") },
            0, 0, OdfRect.FromCentimeters(1, 1, 10, 6)));
        Assert.Empty(chartSheet.Charts);
    }

    [Fact]
    public void AuthoredChartTitleEncodesSignificantWhitespace() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        string title = " Net  Sales\tQ1\nWest";
        sheet.AddChart(OdsChartType.Column, "Data.$A$1:.$A$2",
            new[] { new OdsChartSeries("Data.$B$1:.$B$2") }, 0, 0,
            OdfRect.FromCentimeters(1, 1, 10, 6), title);
        XDocument xml = XDocument.Parse(Encoding.UTF8.GetString(
            document.GetPackageEntryBytes("Object 1/content.xml")));
        Assert.NotEmpty(xml.Descendants(OdfNamespaces.Text + "s"));
        Assert.NotEmpty(xml.Descendants(OdfNamespaces.Text + "tab"));
        Assert.NotEmpty(xml.Descendants(OdfNamespaces.Text + "line-break"));
        OdsDocument reopened = OdsDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(title, Assert.Single(reopened.GetSheet("Data")!.Charts).Title);
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

    [Fact]
    public void VersionRewriteUpdatesEmbeddedChartMetadataAndSettings() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "microsoft-excel-column-chart.ods");
        OdsDocument imported = OdsDocument.Load(path);
        const string office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        foreach (string part in new[] { "meta.xml", "settings.xml" }) {
            string root = part == "meta.xml" ? "document-meta" : "document-settings";
            imported.Package.AddOrReplaceEntry("Object 1/" + part,
                Encoding.UTF8.GetBytes("<office:" + root + " xmlns:office=\"" + office +
                    "\" office:version=\"1.4\"/>"), "text/xml");
        }
        byte[] rewritten = imported.ToBytes(new OdfSaveOptions {
            CompatibilityProfile = OdfCompatibilityProfile.Odf13
        });
        OdsDocument reopened = OdsDocument.Load(new MemoryStream(rewritten));
        foreach (string part in new[] { "meta.xml", "settings.xml" }) {
            XDocument embedded = XDocument.Parse(Encoding.UTF8.GetString(
                reopened.GetPackageEntryBytes("Object 1/" + part)));
            Assert.Equal("1.3", (string?)embedded.Root!.Attribute(OdfNamespaces.Office + "version"));
        }
    }
}
