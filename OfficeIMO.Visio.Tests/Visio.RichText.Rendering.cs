using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioRichTextRenderingTests {
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void IndependentColorAndFontRunsSurviveRenderingAndReopening(int reopen) {
        VisioDocument document = Reopen(Producer("clickhouse-distributed-insert.vdx"), reopen);
        XElement first = ShapeSvg(document.Pages[0], "6");
        Assert.Equal("INSERT INTO ", ColoredText(first, "#009051"));
        Assert.Equal("distributed_table", ColoredText(first, "#000000"));
        XElement second = ShapeSvg(document.Pages[1], "12");
        Assert.Equal("INSERT INTO ", ColoredText(second, "#009051"));
        Assert.Equal("local_table", ColoredText(second, "#000000"));
        XElement fontSwitch = ShapeSvg(document.Pages[1], "10");
        Assert.Contains(fontSwitch.Descendants(Svg + "text"), node =>
            (string?)node.Attribute("font-family") == "Yandex Sans Text" && node.Value.Contains("Асинхронно"));
        Assert.Contains(fontSwitch.Descendants(Svg + "text"), node =>
            (string?)node.Attribute("font-family") == "Monaco" && node.Value.Contains("sharding_key"));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void IndependentEmphasisRunsRetainRegularTextBetweenBoldNames(int reopen) {
        VisioDocument document = Reopen(Producer("nxbre-chocolatebox.vdx"), reopen);
        XElement label = ShapeSvg(document.Pages.Single(page => page.Name == "Rules"), "18");
        string bold = string.Concat(label.Descendants(Svg + "text").Where(node =>
            (string?)node.Attribute("font-weight") == "700").Select(node => node.Value));
        Assert.Equal("org.nxbre.test.ie.ChocolateBoxBinderchocolatebox.vdx.ccb", bold);
        string regular = string.Concat(label.Descendants(Svg + "text").Where(node =>
            node.Attribute("font-weight") == null).Select(node => node.Value));
        Assert.Equal("This implication needs a binder: use , which is stored in ", regular);
        Assert.All(label.Descendants(Svg + "text"), node => Assert.Equal("Arial", (string?)node.Attribute("font-family")));
        Assert.DoesNotContain(label.Descendants(Svg + "rect"), node => node.Attribute("data-officeimo-text-background") != null);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NonzeroRowsAndSourceSectionEditsReachSvgAndRaster(bool connector) {
        VisioDocument document = Synthetic(connector);
        var page = document.Pages[0];
        XElement svg = RenderedLabel(page, connector);
        Assert.Equal("AA", ColoredText(svg, "#cc1122"));
        Assert.Equal("BB", ColoredText(svg, "#008844"));
        XElement first = svg.Descendants(Svg + "text").First();
        Assert.Equal("700", (string?)first.Attribute("font-weight"));
        Assert.Equal("24", (string?)first.Attribute("font-size"));
        Assert.Contains(svg.Descendants(Svg + "text"), node => (string?)node.Attribute("font-style") == "italic");
        XElement raised = svg.Descendants(Svg + "text").Last();
        Assert.Equal("10.4", (string?)raised.Attribute("font-size"));
        Assert.True(double.Parse(raised.Attribute("y")!.Value, System.Globalization.CultureInfo.InvariantCulture)
            < double.Parse(first.Attribute("y")!.Value, System.Globalization.CultureInfo.InvariantCulture));
        Assert.Contains(svg.Descendants(), node => ((string?)node.Attribute("text-decoration") ?? "").Contains("underline"));
        AssertRasterColors(page, red: true, green: true);

        var section = (connector ? page.Connectors[0].GetShapeSheetSections() : page.Shapes[0].GetShapeSheetSections())
            .Single(value => value.Name is "Char" or "Character");
        section.Rows.Single(row => row.Index == 9).Cells.Single(cell => cell.Name == "Color").Value = "#2244cc";
        if (connector) page.Connectors[0].SetShapeSheetSection(section); else page.Shapes[0].SetShapeSheetSection(section);
        foreach (VisioDocument candidate in new[] { document, Reopen(document, 1), Reopen(document, 2) }) {
            Assert.Equal("BB", ColoredText(RenderedLabel(candidate.Pages[0], connector), "#2244cc"));
            Assert.Equal("AA", ColoredText(RenderedLabel(candidate.Pages[0], connector), "#cc1122"));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PlainLabelReplacementDoesNotReuseStaleRunBoundaries(bool connector) {
        VisioDocument document = Synthetic(connector);
        if (connector) document.Pages[0].Connectors[0].Label = "Replacement";
        else document.Pages[0].Shapes[0].Text = "Replacement";
        foreach (VisioDocument candidate in new[] { document, Reopen(document, 1), Reopen(document, 2) }) {
            XElement svg = RenderedLabel(candidate.Pages[0], connector);
            Assert.DoesNotContain(svg.DescendantsAndSelf(), node => node.Attribute("data-officeimo-rich-text") != null);
            Assert.Contains("Replacement", string.Concat(svg.Descendants(Svg + "text").Select(node => node.Value)));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CharacterRunsReachTheSharedShapingProviderAndFontDiagnostics(bool connector) {
        var page = Synthetic(connector).Pages[0];
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new VisioImageExportOptions { TextShapingProvider = provider, Supersampling = 1 };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName, ManagedTextShapingTestAssets.CreateFont('A', 'B'));
        page.ExportImage(OfficeImageExportFormat.Png, options);
        Assert.Contains(provider.Requests, request => request.Text.Contains("AA"));
        Assert.Contains(provider.Requests, request => request.Text.Contains("BB"));
        var export = page.ExportImage(OfficeImageExportFormat.Svg);
        Assert.Contains(export.Diagnostics, diagnostic => diagnostic.Message.Contains(ManagedTextShapingTestAssets.FamilyName));
    }

    [Theory]
    [InlineData("clickhouse-distributed-insert.vdx")]
    [InlineData("nxbre-chocolatebox.vdx")]
    public void RenderingDoesNotRewriteNativeMarkersCellsOrFormulas(string file) {
        VisioDocument document = Producer(file);
        byte[] before = document.ToLegacyXmlResult().Value;
        foreach (VisioPage page in document.Pages) {
            page.ToSvg();
            page.ToPng(new VisioPngSaveOptions { Supersampling = 1 });
            page.ExportImage(OfficeImageExportFormat.Svg);
        }
        Assert.Equal(before, document.ToLegacyXmlResult().Value);
    }

    [Fact]
    public void DuplicatingAPageRetainsMixedRunsWithIndependentSectionEdits() {
        VisioDocument document = Producer("clickhouse-distributed-insert.vdx");
        VisioPage copy = document.DuplicatePage(document.Pages[0], "Copy");
        VisioShape copiedLabel = copy.Shapes.Single(shape => shape.Text == "INSERT INTO distributed_table");
        VisioShape original = document.Pages[0].Shapes.Single(shape => shape.Text == copiedLabel.Text);
        var section = copiedLabel.GetShapeSheetSections().Single(value => value.Name is "Character" or "Char");
        section.Rows.Single(row => row.Index == 0).Cells.Single(cell => cell.Name == "Color").Value = "#2244cc";
        copiedLabel.SetShapeSheetSection(section);
        foreach (VisioDocument candidate in new[] { document, Reopen(document, 1), Reopen(document, 2) }) {
            Assert.Equal("INSERT INTO ", ColoredText(ShapeSvg(candidate.Pages[0], original.Id), "#009051"));
            Assert.Equal("INSERT INTO ", ColoredText(ShapeSvg(candidate.Pages.Single(page => page.Name == "Copy"), copiedLabel.Id), "#2244cc"));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CachedRunTransparencyIsAppliedToItsOwnColor(bool connector) {
        VisioDocument document = Synthetic(connector);
        var page = document.Pages[0];
        var section = (connector ? page.Connectors[0].GetShapeSheetSections() : page.Shapes[0].GetShapeSheetSections())
            .Single(value => value.Name is "Character" or "Char");
        section.Rows.Single(row => row.Index == 7).SetCell("ColorTrans", "0.5");
        if (connector) page.Connectors[0].SetShapeSheetSection(section); else page.Shapes[0].SetShapeSheetSection(section);
        foreach (VisioDocument candidate in new[] { document, Reopen(document, 1), Reopen(document, 2) }) {
            XElement label = RenderedLabel(candidate.Pages[0], connector);
            XElement first = label.Descendants(Svg + "text").First();
            Assert.Equal("0.502", (string?)first.Attribute("fill-opacity"));
            Assert.Null(label.Descendants(Svg + "text").Last().Attribute("fill-opacity"));
        }
    }

    private static VisioDocument Producer(string file) => VisioDocument.LoadLegacyXml(
        Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", file)).Value;

    private static VisioDocument Reopen(VisioDocument document, int mode) => mode switch {
        1 => VisioDocument.Load(new MemoryStream(document.ToBytes())),
        2 => VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value,
        _ => document
    };

    private static XElement ShapeSvg(VisioPage page, string id) => XDocument.Parse(page.ToSvg()).Descendants(Svg + "g")
        .Single(node => (string?)node.Attribute("data-visio-shape-id") == id || (string?)node.Attribute("data-visio-connector-id") == id);

    private static XElement RenderedLabel(VisioPage page, bool connector) => connector
        ? XDocument.Parse(page.ToSvg()).Descendants(Svg + "g").Single(node => node.Attribute("data-visio-connector-id") != null)
        : ShapeSvg(page, "1");

    private static string ColoredText(XElement svg, string color) => string.Concat(svg.Descendants(Svg + "text")
        .Where(node => string.Equals((string?)node.Attribute("fill"), color, StringComparison.OrdinalIgnoreCase)).Select(node => node.Value));

    private static void AssertRasterColors(VisioPage page, bool red, bool green) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ToPng(new VisioPngSaveOptions { Supersampling = 1 }), out var raster));
        byte[] rgba = raster!.GetPixels();
        bool foundRed = false, foundGreen = false;
        for (int offset = 0; offset < rgba.Length; offset += 4) {
            foundRed |= rgba[offset] > 100 && rgba[offset] > rgba[offset + 1] * 2 && rgba[offset] > rgba[offset + 2] * 2;
            foundGreen |= rgba[offset + 1] > 50 && rgba[offset + 1] > rgba[offset] * 2 && rgba[offset + 1] > rgba[offset + 2] * 1.5;
        }
        Assert.Equal(red, foundRed); Assert.Equal(green, foundGreen);
    }

    private static VisioDocument Synthetic(bool connector) {
        string transform = connector
            ? "<XForm1D><BeginX>1</BeginX><BeginY>1</BeginY><EndX>5</EndX><EndY>1</EndY></XForm1D><TextXForm><TxtWidth>4</TxtWidth><TxtHeight>1</TxtHeight></TextXForm>"
            : "<XForm><PinX>3</PinX><PinY>1.5</PinY><Width>4</Width><Height>1</Height></XForm>";
        string source = $"<VisioDocument xmlns='{Legacy}'><Colors><ColorEntry IX='2' RGB='#008844'/></Colors><FaceNames><FaceName ID='0' Name='{ManagedTextShapingTestAssets.FamilyName}'/></FaceNames><Pages><Page ID='0' Name='Runs'><PageSheet><PageProps><PageWidth>6</PageWidth><PageHeight>3</PageHeight></PageProps></PageSheet><Shapes><Shape ID='1'>{transform}<Char IX='7'><Font>0</Font><Color>#cc1122</Color><Style>1</Style><Size Unit='PT'>0.25</Size></Char><Char IX='9'><Font>0</Font><Color>2</Color><Style>6</Style><Size Unit='PT'>0.16666666666666667</Size><DblUnderline>1</DblUnderline><Strikethru>1</Strikethru><Pos>1</Pos></Char><Para IX='4'><HorzAlign>1</HorzAlign><SpLine>-1</SpLine></Para><Text><cp IX='7'/><pp IX='4'/>AA<cp IX='9'/>BB</Text></Shape></Shapes></Page></Pages></VisioDocument>";
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
    }
}
