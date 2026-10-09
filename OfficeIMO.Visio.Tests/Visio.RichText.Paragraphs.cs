using System.Globalization;
using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioRichTextParagraphTests {
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void ProducerParagraphAlignmentAndBlankLineSurviveReopening(int mode) {
        VisioDocument document = Reopen(Producer(), mode);
        VisioPage page = document.Pages.Single(page => page.NameU == "runtime");
        VisioShape shape = page.Shapes.Single(shape => shape.Id == "22");
        XElement[] text = Label(page, "22").Descendants(Svg + "text").ToArray();
        Assert.Equal(new[] { "Event Queue", "(wait)" }, text.Select(node => node.Value));
        double left = (shape.PinX - (shape.TextStyle?.TextWidth ?? shape.Width) / 2D + (shape.TextStyle?.LeftMargin ?? 0D)) * 96D;
        Assert.InRange(Number(text[0], "x") - left, -.002D, .002D);
        Assert.True(Number(text[1], "x") > Number(text[0], "x") + 20D);
        Assert.InRange(Number(text[1], "y") - Number(text[0], "y"), 39.99D, 40.01D);
        Assert.Equal("Event Queue\n\n(wait)", shape.Text);
        Assert.All(text, node => Assert.Equal("700", (string?)node.Attribute("font-weight")));
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(false, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 0)]
    [InlineData(true, 1)]
    [InlineData(true, 2)]
    public void NonzeroParagraphRowsPlaceFirstAndHangingIndentsAndSpacing(bool connector, int mode) {
        VisioDocument document = Reopen(Synthetic(connector), mode);
        XElement[] text = Label(document.Pages[0], "1").Descendants(Svg + "text").ToArray();
        Assert.Equal(new[] { "FIRST", "SECOND", "LAST" }, text.Select(node => node.Value));
        // The second row hangs 1/8 inch inside its 1/4 inch left indent.
        Assert.InRange(Number(text[0], "x") - Number(text[1], "x"), 35.99D, 36.01D);
        Assert.InRange(Number(text[1], "y") - Number(text[0], "y"), 41.99D, 42.01D);
        Assert.InRange(Number(text[2], "y") - Number(text[1], "y"), 29.99D, 30.01D);
        Assert.Equal("16", (string?)text[0].Attribute("font-size"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ParagraphSectionChangesReachBothRenderersAndCopiesRemainIndependent(bool connector) {
        VisioDocument document = Synthetic(connector);
        VisioPage copy = document.DuplicatePage(document.Pages[0], "Paragraph copy");
        byte[] originalRaster = copy.ToPng(new VisioPngSaveOptions { Supersampling = 1 });
        VisioShapeSheetSection section = Sections(document.Pages[0], connector).Single(section => section.Name is "Paragraph" or "Para");
        section.Rows.Single(row => row.Index == 9).SetCell("IndLeft", "0.75");
        if (connector) document.Pages[0].Connectors[0].SetShapeSheetSection(section);
        else document.Pages[0].Shapes[0].SetShapeSheetSection(section);
        string copyId = connector ? copy.Connectors[0].Id : copy.Shapes[0].Id;
        Assert.NotEqual(Number(Label(copy, copyId).Descendants(Svg + "text").ElementAt(1), "x"),
            Number(Label(document.Pages[0], "1").Descendants(Svg + "text").ElementAt(1), "x"));
        Assert.False(originalRaster.SequenceEqual(document.Pages[0].ToPng(new VisioPngSaveOptions { Supersampling = 1 })));
        Assert.Equal(originalRaster, copy.ToPng(new VisioPngSaveOptions { Supersampling = 1 }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReplacingPlainTextDiscardsNativeParagraphPlacement(bool connector) {
        VisioDocument document = Synthetic(connector);
        if (connector) document.Pages[0].Connectors[0].Label = "Replacement";
        else document.Pages[0].Shapes[0].Text = "Replacement";
        foreach (VisioDocument candidate in new[] { document, Reopen(document, 1), Reopen(document, 2) }) {
            XElement label = Label(candidate.Pages[0], "1");
            Assert.Equal("Replacement", string.Concat(label.Descendants(Svg + "text").Select(node => node.Value)));
            Assert.DoesNotContain(label.Descendants(Svg + "g"), node => node.Attribute("data-officeimo-rich-text") != null);
        }
    }

    [Fact]
    public void RenderingProducerParagraphsDoesNotChangeNativeCellsOrMarkers() {
        VisioDocument document = Producer();
        byte[] before = document.ToLegacyXmlResult().Value;
        VisioPage page = document.Pages.Single(page => page.NameU == "runtime");
        _ = page.ToSvg(); _ = page.ToPng(new VisioPngSaveOptions { Supersampling = 1 });
        Assert.Equal(before, document.ToLegacyXmlResult().Value);
    }

    [Fact]
    public void ParagraphFittingPreservesContentAndKeepsNativeInsets() {
        VisioDocument document = Synthetic(false, height: .8D);
        XElement[] text = Label(document.Pages[0], "1").Descendants(Svg + "text").ToArray();
        Assert.Equal(new[] { "FIRST", "SECOND", "LAST" }, text.Select(node => node.Value));
        Assert.InRange(Number(text[0], "font-size"), 5D, 15.99D);
        Assert.InRange(Number(text[0], "x") - Number(text[1], "x"), 35.99D, 36.01D);
    }

    [Fact]
    public void WrappedContinuationUsesHangingIndentWithoutAddingParagraphSpacing() {
        VisioDocument document = Synthetic(false, second: "SECOND ONE TWO THREE FOUR FIVE SIX SEVEN EIGHT NINE TEN ELEVEN TWELVE THIRTEEN");
        XElement[] text = Label(document.Pages[0], "1").Descendants(Svg + "text").ToArray();
        Assert.True(text.Length > 3);
        Assert.StartsWith("SECOND", text[1].Value);
        Assert.InRange(Number(text[2], "x") - Number(text[1], "x"), 11.99D, 12.01D);
        Assert.InRange(Number(text[2], "y") - Number(text[1], "y"), 23.99D, 24.01D);
    }

    [Fact]
    public void CancellationDuringParagraphMeasurementStopsBeforeLaterParagraphs() {
        using var cancellation = new CancellationTokenSource();
        var provider = new CancelProvider(cancellation);
        var options = new VisioImageExportOptions { TextShapingProvider = provider, Supersampling = 1 };
        options.Fonts.Add("Arial", ManagedTextShapingTestAssets.CreateFont('F', 'I', 'R', 'S', 'T', 'E', 'C', 'O', 'N', 'D', 'L', 'A'));
        Assert.Throws<OperationCanceledException>(() => Synthetic(false).Pages[0].ExportImage(
            OfficeImageExportFormat.Png, options, cancellation.Token));
        Assert.InRange(provider.Requests, 1, 2);
    }

    private static VisioDocument Producer() => VisioDocument.LoadLegacyXml(Path.Combine(AppContext.BaseDirectory,
        "Fixtures", "LegacyXml", "angularjs-concepts.vdx")).Value;

    private static VisioDocument Reopen(VisioDocument document, int mode) => mode switch {
        1 => VisioDocument.Load(new MemoryStream(document.ToBytes())),
        2 => VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value,
        _ => document
    };

    private static IReadOnlyList<VisioShapeSheetSection> Sections(VisioPage page, bool connector) => connector
        ? page.Connectors[0].GetShapeSheetSections() : page.Shapes[0].GetShapeSheetSections();

    private static XElement Label(VisioPage page, string id) => XDocument.Parse(page.ToSvg()).Descendants(Svg + "g")
        .Single(node => (string?)node.Attribute("data-visio-shape-id") == id || (string?)node.Attribute("data-visio-connector-id") == id);

    private static double Number(XElement node, string attribute) => double.Parse(node.Attribute(attribute)!.Value, CultureInfo.InvariantCulture);

    private sealed class CancelProvider(CancellationTokenSource cancellation) : IOfficeTextShapingProvider {
        internal int Requests { get; private set; }
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            Requests++; cancellation.Cancel();
            return new ManagedTextShapingTestAssets.RecordingProvider().ShapeText(request);
        }
    }

    private static VisioDocument Synthetic(bool connector, double height = 2D, string second = "SECOND") {
        string transform = connector
            ? $"<XForm1D><BeginX>1</BeginX><BeginY>2</BeginY><EndX>5</EndX><EndY>2</EndY></XForm1D><TextXForm><TxtWidth>4</TxtWidth><TxtHeight>{height.ToString(CultureInfo.InvariantCulture)}</TxtHeight></TextXForm>"
            : $"<XForm><PinX>3</PinX><PinY>2</PinY><Width>4</Width><Height>{height.ToString(CultureInfo.InvariantCulture)}</Height></XForm>";
        string source = $"<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'><FaceNames><FaceName ID='0' Name='Arial'/></FaceNames><Pages><Page ID='0' Name='Paragraphs'><PageSheet><PageProps><PageWidth>6</PageWidth><PageHeight>4</PageHeight></PageProps></PageSheet><Shapes><Shape ID='1'>{transform}<TextBlock><VerticalAlign>0</VerticalAlign></TextBlock><Char IX='3'><Font>0</Font><Size>0.16666666666666667</Size><Color>#cc1122</Color></Char><Para IX='7'><IndFirst>0.25</IndFirst><IndLeft>0.25</IndLeft><IndRight>0.125</IndRight><SpLine>0.25</SpLine><SpAfter>0.125</SpAfter><HorzAlign>0</HorzAlign></Para><Para IX='9'><IndFirst>-0.125</IndFirst><IndLeft>0.25</IndLeft><IndRight>0.125</IndRight><SpLine>-1.5</SpLine><SpBefore>0.0625</SpBefore><HorzAlign>0</HorzAlign></Para><Text><cp IX='3'/><pp IX='7'/>FIRST\n<pp IX='9'/>{second}\nLAST</Text></Shape></Shapes></Page></Pages></VisioDocument>";
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
    }
}
