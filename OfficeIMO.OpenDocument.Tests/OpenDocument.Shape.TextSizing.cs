using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument.Testing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentShapeTextSizingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FlagsAndInstanceSizesRoundTripInDrawAndPresentation(bool flat) {
        var drawing = OdgDocument.Create(); var box = AddBox(drawing.AddPage());
        ConfigureSizes(box);
        using var drawStream = new MemoryStream();
        if (flat) drawing.SaveFlatXml(drawStream); else drawing.Save(drawStream);
        drawStream.Position = 0;
        AssertSizes((flat ? OdgDocument.LoadFlatXml(drawStream) : OdgDocument.Load(drawStream)).Pages[0].Shapes[0]);
        var presentation = OdpPresentation.Create();
        var presentationBox = presentation.AddSlide().AddTextBox(Bounds(), "Body", "Sizing");
        ConfigureSizes(presentationBox);
        using var presentationStream = new MemoryStream();
        if (flat) presentation.SaveFlatXml(presentationStream); else presentation.Save(presentationStream);
        presentationStream.Position = 0;
        AssertSizes((flat ? OdpPresentation.LoadFlatXml(presentationStream) : OdpPresentation.Load(presentationStream)).Slides[0].Shapes[0]);
    }

    [Fact]
    public void GrowthAxesInheritIndependentlyAndCopyOnWriteKeepsSharedShapesIsolated() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var first = AddBox(page); var second = AddBox(page);
        var parent = document.Styles.CreateNamed("GrowthParent", OdfStyleFamily.Graphic);
        Set(parent, "auto-grow-height", "true"); Set(parent, "auto-grow-width", "false");
        var shared = document.Styles.CreateAutomatic(OdfStyleFamily.Graphic, parentStyleName: parent.Name);
        Set(shared, "auto-grow-height", "false"); Set(shared, "auto-grow-width", "true");
        Bind(first, shared); Bind(second, shared);
        first.AutoGrowHeight = true;
        Assert.True(first.AutoGrowHeight); Assert.True(first.AutoGrowWidth);
        Assert.False(second.AutoGrowHeight); Assert.True(second.AutoGrowWidth);
        first.AutoGrowWidth = null;
        Assert.False(first.AutoGrowWidth); Assert.True(second.AutoGrowWidth);
        first.AutoGrowHeight = null;
        Assert.True(first.AutoGrowHeight);
        Assert.NotEqual((string?)first.Element.Attribute(OdfNamespaces.Draw + "style-name"),
            (string?)second.Element.Attribute(OdfNamespaces.Draw + "style-name"));
        first.TextBoxMinimumHeight = OdfLength.Parse("24pt");
        first.TextBoxMaximumHeight = OdfLength.Parse("96pt");
        first.TextBoxMinimumHeight = null; first.TextBoxMaximumHeight = null;
        Assert.Null(first.TextBoxMinimumHeight); Assert.Null(first.TextBoxMaximumHeight);
        Assert.Null(second.TextBoxMinimumHeight);
    }

    [Theory]
    [InlineData("1")]
    [InlineData("grow")]
    public void UnknownGrowthIsPreservedAndReportedWithoutGetterOrProjectionMutation(string lexical) {
        var document = OdgDocument.Create(); var shape = AddBox(document.AddPage());
        Set(Style(shape), "auto-grow-height", lexical);
        string[] before = XmlState(document);
        Assert.Throws<NotSupportedException>(() => shape.AutoGrowHeight);
        Assert.Equal(before, XmlState(document));
        var result = document.Pages[0].ToDrawing();
        AssertLoss(result.Report, "text-auto-size");
        Assert.Throws<OdfConversionLossException>(() => document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal(before, XmlState(document));
        shape.AutoGrowHeight = false;
        Assert.False(shape.AutoGrowHeight);
    }

    [Theory]
    [InlineData("-1pt")]
    [InlineData("NaNpt")]
    [InlineData("1e100pt")]
    [InlineData("none")]
    public void InvalidSizeAssignmentsAndWrongContainersLeaveXmlUnchanged(string lexical) {
        var document = OdgDocument.Create(); var page = document.AddPage(); var box = AddBox(page);
        var rectangle = page.Shapes.AddRectangle(Bounds(), "Rectangle");
        string[] before = XmlState(document);
        Assert.Throws<ArgumentException>(() => box.TextBoxMinimumHeight = OdfLength.Parse(lexical));
        Assert.Throws<NotSupportedException>(() => rectangle.TextBoxMaximumWidth = OdfLength.Points(20));
        Assert.Throws<NotSupportedException>(() => rectangle.TextBoxMinimumWidth);
        Assert.Equal(before, XmlState(document));
    }

    [Theory]
    [InlineData(OfficeTextAreaAlignment.FullWidth)]
    [InlineData(OfficeTextAreaAlignment.Left)]
    [InlineData(OfficeTextAreaAlignment.Center)]
    [InlineData(OfficeTextAreaAlignment.Right)]
    public void HeightGrowthResizesPaintAndTextTogetherWithoutChangingSourceOrFontSize(OfficeTextAreaAlignment area) {
        var document = OdgDocument.Create(); var page = document.AddPage(); var shape = AddBox(page);
        shape.Text = string.Join(" ", Enumerable.Repeat("Alpha beta gamma delta", 12)) + " END_MARKER";
        shape.Paragraphs[0].FontSize = OdfLength.Points(12);
        shape.FillColor = OdfColor.Parse("#e0f0ff"); shape.StrokeColor = OdfColor.Parse("#204060");
        shape.AutoGrowHeight = true; shape.AutoGrowWidth = false;
        shape.TextAreaAlignment = area; shape.WrapText = true; shape.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Top;
        shape.TextPadding = new OdfInsets(OdfLength.Points(2), OdfLength.Points(2), OdfLength.Points(2), OdfLength.Points(2));
        string[] before = XmlState(document);
        var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
        var paint = Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>());
        Assert.True(text.Height > 48); Assert.Equal(120, text.Width);
        Assert.Equal(text.Height, paint.Shape.Height); Assert.Equal(text.X, paint.X); Assert.Equal(text.Y, paint.Y);
        Assert.IsType<OfficeDrawingShape>(result.Value.Elements[0]);
        Assert.All(text.Paragraphs[0].Runs, run => Assert.Equal(12, run.FontSize));
        Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":text-auto-size", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Approximated);
        Assert.Contains("END_MARKER", string.Concat(XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(result.Value))
            .Descendants(XNamespace.Get("http://www.w3.org/2000/svg") + "text").Select(e => e.Value)));
        Assert.Equal(before, XmlState(document));

        // A producer-saved grown height remains the minimum after later text removal.
        shape.Bounds = new OdfRect(shape.Bounds.X, shape.Bounds.Y, shape.Bounds.Width, OdfLength.Points(text.Height));
        shape.Text = "Short"; shape.Paragraphs[0].FontSize = OdfLength.Points(12);
        var shortened = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        Assert.Equal(shape.Bounds.Height.ToPoints(), Assert.Single(shortened.Value.Elements.OfType<OfficeDrawingRichText>()).Height);
    }

    [Fact]
    public void SuppliedFontContextChangesGrowthAndPreservesTheCompleteRenderedBody() {
        var document = OdgDocument.Create(); var page = document.AddPage(); var shape = AddBox(page);
        shape.Bounds = new OdfRect(OdfLength.Points(20), OdfLength.Points(20), OdfLength.Points(40), OdfLength.Points(18));
        shape.Text = OdfTextFittingTestFonts.Body;
        shape.Paragraphs[0].FontSize = OdfLength.Points(12); shape.Paragraphs[0].FontFamily = OdfTextFittingTestFonts.Family;
        shape.AutoGrowHeight = true; shape.AutoGrowWidth = false; shape.WrapText = true;
        double[] heights = new double[2];
        foreach (bool wide in new[] { false, true }) {
            var result = page.ToDrawing(OdfTextFittingTestFonts.Profile(wide), OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            heights[wide ? 1 : 0] = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).Height;
            string rendered = string.Concat(XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(result.Value))
                .Descendants(XNamespace.Get("http://www.w3.org/2000/svg") + "text").Select(e => e.Value));
            Assert.Equal(OdfTextFittingTestFonts.Body, rendered);
            Assert.Equal(OdfTextFittingTestFonts.Create(wide), Assert.Single(result.Value.Fonts.Faces).Data);
        }
        Assert.True(heights[1] > heights[0]);
    }

    [Theory]
    [InlineData("unspecified-width")]
    [InlineData("middle")]
    [InlineData("fitting")]
    [InlineData("relative")]
    [InlineData("constraint")]
    public void UnsupportedGrowthProfilesRetainSavedHeightAndFailStrictProjection(string profile) {
        var document = OdgDocument.Create(); var page = document.AddPage(); var shape = AddBox(page);
        shape.AutoGrowHeight = true; shape.AutoGrowWidth = false; shape.WrapText = true;
        switch (profile) {
            case "unspecified-width": shape.AutoGrowWidth = null; break;
            case "middle": shape.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Middle; break;
            case "fitting": shape.TextFitMode = OdfTextFitMode.ShrinkToFit; break;
            case "relative": shape.Element.SetAttributeValue(OdfNamespaces.Style + "rel-height", "100%"); document.MarkPartDirty("content.xml"); break;
            case "constraint": shape.TextBoxMaximumHeight = OdfLength.Points(96); break;
        }
        string[] before = XmlState(document); var result = page.ToDrawing();
        Assert.Equal(48, Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).Height);
        AssertLoss(result.Report, "text-auto-size");
        if (profile == "constraint") AssertLoss(result.Report, "text-size-constraints");
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal(before, XmlState(document));
    }

    [Fact]
    public void InstanceConstraintsAreReportedEvenWithoutTextAndCreationDefaultsDoNotResizeSavedFrames() {
        var document = OdgDocument.Create(); var page = document.AddPage(); var box = AddBox(page);
        Style(box).SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "min-height", "72pt");
        var ordinary = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        Assert.Equal(48, Assert.Single(ordinary.Value.Elements.OfType<OfficeDrawingRichText>()).Height);
        Assert.Null(box.TextBoxMinimumHeight);
        box.Text = ""; box.TextBoxMinimumHeight = OdfLength.Parse("72pt");
        string[] before = XmlState(document); AssertLoss(page.ToDrawing().Report, "text-size-constraints");
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal(before, XmlState(document));
    }

    private static OdgShape AddBox(OdgPage page) {
        var shape = page.Shapes.AddTextBox(Bounds(), "Body", "Sizing");
        shape.Paragraphs[0].FontSize = OdfLength.Points(12);
        return shape;
    }
    private static OdfRect Bounds() => new OdfRect(OdfLength.Points(20), OdfLength.Points(20), OdfLength.Points(120), OdfLength.Points(48));
    private static void ConfigureSizes(OdfShape shape) {
        shape.AutoGrowHeight = true; shape.AutoGrowWidth = false;
        shape.TextBoxMinimumHeight = OdfLength.Parse("24.00pt"); shape.TextBoxMaximumHeight = OdfLength.Parse("96pt");
        shape.TextBoxMinimumWidth = OdfLength.Parse("25%"); shape.TextBoxMaximumWidth = OdfLength.Parse("80px");
    }
    private static void AssertSizes(OdfShape shape) {
        Assert.True(shape.AutoGrowHeight); Assert.False(shape.AutoGrowWidth);
        Assert.Equal("24.00pt", shape.TextBoxMinimumHeight?.ToString()); Assert.Equal("96pt", shape.TextBoxMaximumHeight?.ToString());
        Assert.Equal("25%", shape.TextBoxMinimumWidth?.ToString()); Assert.Equal("80px", shape.TextBoxMaximumWidth?.ToString());
        Assert.Equal(Bounds(), shape.Bounds);
    }
    private static OdfStyle Style(OdfShape shape) => shape.Document.Styles.FindInPart(OdfStyleFamily.Graphic,
        (string)shape.Element.Attribute(OdfNamespaces.Draw + "style-name")!, shape.PartPath)!;
    private static void Set(OdfStyle style, string name, string value) => style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + name, value);
    private static void Bind(OdfShape shape, OdfStyle style) {
        shape.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", style.Name); shape.Document.MarkPartDirty(shape.PartPath);
    }
    private static string[] XmlState(OdfDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString() };
    private static void AssertLoss(OdfConversionReport report, string loss) => Assert.Contains(report.Mappings,
        m => m.Feature.EndsWith(":" + loss, StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
}
