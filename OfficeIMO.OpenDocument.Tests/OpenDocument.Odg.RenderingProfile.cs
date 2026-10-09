using System;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument.Testing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgRenderingProfileTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SuppliedFontWidthsGovernClippingReportsAndTheRenderedContent(bool wide) {
        var document = CreateDocument(); var page = document.Pages[0];
        OfficeRenderingProfile profile = OdfTextFittingTestFonts.Profile(wide);
        string[] before = XmlState(document); var result = page.ToDrawing(profile);
        bool clipped = result.Report.Mappings.Any(m => m.Feature.EndsWith(":text-clipped", StringComparison.Ordinal));
        Assert.Equal(wide, clipped);
        string rendered = string.Concat(XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(result.Value))
            .Descendants(XNamespace.Get("http://www.w3.org/2000/svg") + "text").Select(text => text.Value));
        Assert.Equal(!wide, rendered == OdfTextFittingTestFonts.Body);
        if (wide) Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(profile, OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        else page.ToDrawing(profile, OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        Assert.Equal(before, XmlState(document));
    }

    [Fact]
    public void ProfileSnapshotsReachMasterNestedAndFinalDrawingResources() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        Configure(page.MasterShapes.AddTextBox(Bounds(), OdfTextFittingTestFonts.Body, "Master"));
        var group = page.Shapes.AddGroup("Nested");
        var child = group.Children.AddTextBox(Bounds(), OdfTextFittingTestFonts.Body, "Child");
        Configure(child); child.Transform = "rotate(0.1)";
        var fonts = new OfficeFontFaceCollection().Add(OdfTextFittingTestFonts.Family, OdfTextFittingTestFonts.Create(false));
        var provider = new RecordingProvider();
        var profile = new OfficeRenderingProfile("snapshot", fonts, provider, "en-US");
        fonts.Add(OdfTextFittingTestFonts.Family, OdfTextFittingTestFonts.Create(true));
        string[] before = XmlState(document);
        var result = document.ToDrawings(profile, OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        var drawing = Assert.Single(result.Value); AssertResources(drawing);
        Assert.Single(drawing.Elements.OfType<OfficeDrawingRichText>());
        var nested = Assert.Single(drawing.Elements.OfType<OfficeDrawingEffectGroup>());
        OfficeDrawing nestedDrawing = nested.Drawing;
        AssertResources(nestedDrawing);
        Assert.Equal(OdfTextFittingTestFonts.Body, Assert.Single(nestedDrawing.Elements.OfType<OfficeDrawingRichText>()).PlainText);
        Assert.Equal(before, XmlState(document));
        provider.Languages.Clear();
        OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.NotEmpty(provider.Languages);
        Assert.All(provider.Languages, language => Assert.Equal("en-US", language));

        void AssertResources(OfficeDrawing actual) {
            Assert.Equal(OdfTextFittingTestFonts.Create(false), Assert.Single(actual.Fonts.Faces).Data);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SuppliedShapingPreflightObservesTheOperationCancellationToken(bool growing) {
        var document = CreateDocument(); using var cancellation = new CancellationTokenSource();
        document.Pages[0].Shapes[0].AutoGrowHeight = growing;
        document.Pages[0].Shapes[0].AutoGrowWidth = false;
        var provider = new CancellingProvider(cancellation);
        var profile = OdfTextFittingTestFonts.Profile(false, provider);
        string[] before = XmlState(document);
        Assert.ThrowsAny<OperationCanceledException>(() => document.ToDrawings(profile, cancellationToken: cancellation.Token));
        Assert.Equal(cancellation.Token, provider.ObservedToken); Assert.Equal(before, XmlState(document));
    }

    private sealed class RecordingProvider : IOfficeTextShapingProvider {
        internal System.Collections.Generic.List<string?> Languages { get; } = new();
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            Languages.Add(request.Language);
            return OfficeManagedTextShapingProvider.Instance.ShapeText(request);
        }
    }

    private sealed class CancellingProvider(CancellationTokenSource source) : IOfficeTextShapingProvider {
        internal CancellationToken ObservedToken { get; private set; }
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            ObservedToken = request.CancellationToken; source.Cancel(); request.CancellationToken.ThrowIfCancellationRequested();
            return null;
        }
    }
    private static OdgDocument CreateDocument() {
        var document = OdgDocument.Create();
        Configure(document.AddPage().Shapes.AddTextBox(Bounds(), OdfTextFittingTestFonts.Body, "Body"));
        return document;
    }
    private static OdfRect Bounds() => new OdfRect(OdfLength.Points(20), OdfLength.Points(30), OdfLength.Points(40), OdfLength.Points(20));
    private static void Configure(OdgShape shape) {
        shape.Paragraphs[0].FontFamily = OdfTextFittingTestFonts.Family; shape.Paragraphs[0].FontSize = OdfLength.Points(12);
        shape.WrapText = true; shape.TextPadding = new OdfInsets(OdfLength.Points(0), OdfLength.Points(0), OdfLength.Points(0), OdfLength.Points(0));
    }
    private static string[] XmlState(OdfDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString() };
}
