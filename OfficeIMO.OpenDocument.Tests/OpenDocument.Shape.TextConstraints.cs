using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentShapeTextConstraintsTests {
    [Theory]
    [InlineData(false, 60, 24, false)]
    [InlineData(true, 180, 72, false)]
    [InlineData(false, 60, 24, true)]
    [InlineData(true, 180, 72, true)]
    public void InstanceMinimaReplaceSavedDimensionsForTextAndPaintAfterRoundTrip(bool flat, double width, double height, bool empty) {
        var document = Create(empty ? "" : "Body"); var box = document.Pages[0].Shapes[0];
        box.TextBoxMinimumWidth = OdfLength.Points(width); box.TextBoxMaximumWidth = OdfLength.Points(width + 20);
        box.TextBoxMinimumHeight = OdfLength.Points(height); box.TextBoxMaximumHeight = OdfLength.Points(height + 20);
        using var stream = new MemoryStream();
        if (flat) document.SaveFlatXml(stream); else document.Save(stream);
        stream.Position = 0; var loaded = flat ? OdgDocument.LoadFlatXml(stream) : OdgDocument.Load(stream);
        string[] before = XmlState(loaded);
        var result = loaded.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        var paint = Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>());
        Assert.Equal(width, paint.Shape.Width); Assert.Equal(height, paint.Shape.Height);
        Assert.Equal(20, paint.X); Assert.Equal(30, paint.Y);
        if (empty) Assert.Empty(result.Value.Elements.OfType<OfficeDrawingRichText>());
        else {
            var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
            Assert.Equal(width, text.Width); Assert.Equal(height, text.Height);
            Assert.Equal(paint.X, text.X); Assert.Equal(paint.Y, text.Y);
        }
        Assert.Equal(120, loaded.Pages[0].Shapes[0].Bounds.Width.ToPoints());
        Assert.Equal(48, loaded.Pages[0].Shapes[0].Bounds.Height.ToPoints());
        Assert.Equal(before, XmlState(loaded));
    }

    [Fact]
    public void HeightMinimumReplacesSavedFloorAndMaximumCapsMeasuredGrowthWithReportedClipping() {
        var document = Create("Body"); var page = document.Pages[0]; var box = page.Shapes[0];
        box.AutoGrowHeight = true; box.TextBoxMinimumHeight = OdfLength.Points(24);
        Assert.Equal(24, Paint(page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported)).Shape.Height);
        box.TextBoxMinimumHeight = OdfLength.Points(0);
        double shortHeight = Paint(page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported)).Shape.Height;
        Assert.InRange(shortHeight, 1, 23.999);
        box.Text = string.Join(" ", Enumerable.Repeat("Alpha beta gamma delta", 12));
        box.Paragraphs[0].FontSize = OdfLength.Points(12);
        box.TextBoxMinimumHeight = OdfLength.Points(24); box.TextBoxMaximumHeight = OdfLength.Points(96);
        string[] before = XmlState(document); var capped = page.ToDrawing();
        Assert.Equal(96, Paint(capped).Shape.Height);
        Assert.Equal(96, Assert.Single(capped.Value.Elements.OfType<OfficeDrawingRichText>()).Height);
        AssertLoss(capped.Report, "text-clipped");
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal(before, XmlState(document));
        box.TextBoxMaximumHeight = OdfLength.Points(1000);
        double grown = Paint(page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported)).Shape.Height;
        Assert.InRange(grown, 96.001, 999.999);
        box.TextBoxMaximumHeight = null;
        Assert.Equal(grown, Paint(page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported)).Shape.Height);
    }

    [Theory]
    [InlineData("orphan")]
    [InlineData("units")]
    [InlineData("inverted")]
    [InlineData("relative")]
    [InlineData("pixels")]
    public void UnsupportedPairsArePreservedAndDoNotPartiallyResizeOtherAxis(string profile) {
        var document = Create(""); var page = document.Pages[0]; var box = page.Shapes[0];
        box.TextBoxMinimumWidth = OdfLength.Points(60);
        switch (profile) {
            case "orphan": box.TextBoxMaximumHeight = OdfLength.Points(96); break;
            case "units": box.TextBoxMinimumHeight = OdfLength.Points(24); box.TextBoxMaximumHeight = OdfLength.Parse("2in"); break;
            case "inverted": box.TextBoxMinimumHeight = OdfLength.Points(96); box.TextBoxMaximumHeight = OdfLength.Points(24); break;
            case "relative": box.TextBoxMinimumHeight = OdfLength.Parse("50%"); break;
            case "pixels": box.TextBoxMinimumHeight = OdfLength.Parse("24px"); break;
        }
        string[] before = XmlState(document); var result = page.ToDrawing();
        Assert.Equal(120, Paint(result).Shape.Width); Assert.Equal(48, Paint(result).Shape.Height);
        AssertLoss(result.Report, "text-size-constraints");
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal(before, XmlState(document));
    }

    [Theory]
    [InlineData(false, 0, 48)]
    [InlineData(true, 120, 0)]
    public void ZeroFrameOmitsPaintAndNonemptyBodyReportsClipping(bool empty, double width, double height) {
        var document = Create(empty ? "" : "Body"); var page = document.Pages[0]; var box = page.Shapes[0];
        box.TextBoxMinimumWidth = OdfLength.Points(width); box.TextBoxMinimumHeight = OdfLength.Points(height);
        string[] before = XmlState(document); var result = page.ToDrawing();
        Assert.Empty(result.Value.Elements.OfType<OfficeDrawingShape>());
        Assert.Empty(result.Value.Elements.OfType<OfficeDrawingRichText>());
        if (empty) page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        else {
            AssertLoss(result.Report, "text-clipped");
            Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        }
        Assert.Equal(before, XmlState(document));
    }

    [Fact]
    public void ZeroMaximumStopsHeightGrowthAndReportsOmittedBody() {
        var document = Create("Body"); var page = document.Pages[0]; var box = page.Shapes[0];
        box.AutoGrowHeight = true; box.TextBoxMinimumHeight = OdfLength.Points(0); box.TextBoxMaximumHeight = OdfLength.Points(0);
        string[] before = XmlState(document); var result = page.ToDrawing();
        Assert.Empty(result.Value.Elements.OfType<OfficeDrawingShape>());
        Assert.Empty(result.Value.Elements.OfType<OfficeDrawingRichText>());
        AssertLoss(result.Report, "text-clipped");
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal(before, XmlState(document));
    }

    private static OdgDocument Create(string body) {
        var document = OdgDocument.Create(); var box = document.AddPage().Shapes.AddTextBox(new OdfRect(
            OdfLength.Points(20), OdfLength.Points(30), OdfLength.Points(120), OdfLength.Points(48)), body, "Constraints");
        box.AutoGrowWidth = false; box.AutoGrowHeight = false; box.WrapText = true;
        box.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Top;
        box.TextPadding = new OdfInsets(OdfLength.Points(0), OdfLength.Points(0), OdfLength.Points(0), OdfLength.Points(0));
        box.FillColor = OdfColor.Parse("#e0f0ff"); box.StrokeColor = OdfColor.Parse("#204060");
        if (!string.IsNullOrEmpty(body)) box.Paragraphs[0].FontSize = OdfLength.Points(12);
        return document;
    }
    private static OfficeDrawingShape Paint(OdfConversionResult<OfficeDrawing> result) => Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>());
    private static string[] XmlState(OdfDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString() };
    private static void AssertLoss(OdfConversionReport report, string name) => Assert.Contains(report.Mappings,
        m => m.Feature.EndsWith(":" + name, StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
}
