using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentShapeTextFittingTests {
    [Theory]
    [InlineData(OdfTextFitMode.None, "false", "false")]
    [InlineData(OdfTextFitMode.Stretch, "true", "false")]
    [InlineData(OdfTextFitMode.ShrinkToFit, "false", "true")]
    public void CanonicalModesPersistInDrawingFlatDrawingAndPresentation(OdfTextFitMode mode, string stretch, string shrink) {
        var drawing = OdgDocument.Create(); var shape = AddBox(drawing.AddPage(), "Fitting");
        shape.TextFitMode = mode;
        foreach (bool flat in new[] { false, true }) {
            var actual = Reopen(drawing, flat).Pages[0].Shapes[0];
            AssertMode(actual, mode, stretch, shrink); Assert.Equal(Bounds(), actual.Bounds);
        }
        var presentation = OdpPresentation.Create(); var box = presentation.AddSlide().AddTextBox(Bounds(), "Presentation", "Fitting");
        box.TextFitMode = mode;
        var reopened = OdpPresentation.Load(new MemoryStream(presentation.ToBytes()));
        AssertMode(Assert.IsType<OdpTextBox>(Assert.Single(reopened.Slides[0].Shapes)), mode, stretch, shrink);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EditingSharedFittingIsCopyOnWriteAndNullRestoresTheNamedParent(bool flat) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var first = AddBox(page, "First"); var second = AddBox(page, "Second");
        var parent = document.Styles.CreateNamed("StretchParent", OdfStyleFamily.Graphic);
        SetPair(parent, "true", "false");
        var shared = document.Styles.CreateAutomatic(OdfStyleFamily.Graphic, parentStyleName: parent.Name);
        SetPair(shared, "false", "true");
        Bind(first, shared); Bind(second, shared);
        document = Reopen(document, flat); first = document.Pages[0].Shapes[0]; second = document.Pages[0].Shapes[1];
        string[] before = XmlState(document);
        Assert.Equal(OdfTextFitMode.ShrinkToFit, first.TextFitMode); Assert.Equal(before, XmlState(document));
        Assert.Throws<ArgumentOutOfRangeException>(() => first.TextFitMode = (OdfTextFitMode)99);
        Assert.Equal(before, XmlState(document));
        first.TextFitMode = OdfTextFitMode.None;
        Assert.NotEqual(GraphicStyle(first).Name, GraphicStyle(second).Name);
        AssertMode(first, OdfTextFitMode.None, "false", "false");
        AssertMode(second, OdfTextFitMode.ShrinkToFit, "false", "true");
        first.TextFitMode = null;
        var actual = Reopen(document, flat);
        Assert.Equal(OdfTextFitMode.Stretch, actual.Pages[0].Shapes[0].TextFitMode);
        AssertPair(actual.Pages[0].Shapes[0], null, null);
        Assert.Equal(OdfTextFitMode.ShrinkToFit, actual.Pages[0].Shapes[1].TextFitMode);
    }

    [Theory]
    [InlineData("false", null, "false", "true")]
    [InlineData(null, "false", "true", "false")]
    public void EitherNearestDeclarationOwnsTheModeAndSettingShrinkReplacesInheritedStretch(
        string? localStretch, string? localShrink, string parentStretch, string parentShrink) {
        var document = OdgDocument.Create(); var shape = AddBox(document.AddPage(), "Nearest");
        Assert.Null(shape.TextFitMode);
        var parent = document.Styles.CreateNamed("Parent", OdfStyleFamily.Graphic); SetPair(parent, parentStretch, parentShrink);
        var local = document.Styles.CreateAutomatic(OdfStyleFamily.Graphic, parentStyleName: parent.Name);
        SetPair(local, localStretch, localShrink); Bind(shape, local);
        string[] before = XmlState(document);
        Assert.Equal(OdfTextFitMode.None, shape.TextFitMode); Assert.Equal(before, XmlState(document));
        shape.TextFitMode = null;
        Assert.Equal(parentStretch == "true" ? OdfTextFitMode.Stretch : OdfTextFitMode.ShrinkToFit, shape.TextFitMode);
        shape.TextFitMode = OdfTextFitMode.ShrinkToFit;
        AssertMode(shape, OdfTextFitMode.ShrinkToFit, "false", "true");
        Assert.Equal(parentStretch, (string?)parent.Element.Element(OdfNamespaces.Style + "graphic-properties")!.Attribute(OdfNamespaces.Draw + "fit-to-size"));
        Assert.Equal(parentShrink, (string?)parent.Element.Element(OdfNamespaces.Style + "graphic-properties")!.Attribute(OdfNamespaces.Style + "shrink-to-fit"));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void PageAndMasterImportsKeepAnExplicitPairMemberAboveOppositeDefaults(bool master, bool named) {
        foreach (bool defaultShrink in new[] { false, true }) {
            var source = OdgDocument.Create(); var page = source.AddPage("Source");
            var shape = (master ? page.MasterShapes : page.Shapes).AddTextBox(Bounds(), "Local fitting", "Local");
            OdfStyle style = named ? source.Styles.CreateNamed("Local", OdfStyleFamily.Graphic) : GraphicStyle(shape);
            if (named) Bind(shape, style);
            // Native nearest-pair semantics: one local false overrides the entire default opposite mode.
            SetPair(style, defaultShrink ? "false" : null, defaultShrink ? null : "false");
            source.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(
                new XElement(OdfNamespaces.Style + "default-style", new XAttribute(OdfNamespaces.Style + "family", "graphic"),
                    new XElement(OdfNamespaces.Style + "graphic-properties",
                        new XAttribute(OdfNamespaces.Draw + "fit-to-size", defaultShrink ? "false" : "true"),
                        new XAttribute(OdfNamespaces.Style + "shrink-to-fit", defaultShrink ? "true" : "false"),
                        new XAttribute(OdfNamespaces.Fo + "wrap-option", "no-wrap"))));
            source.MarkPartDirty("styles.xml");
            foreach (bool flat in new[] { false, true }) {
                var original = Reopen(source, flat);
                var originalShape = (master ? original.Pages[0].MasterShapes : original.Pages[0].Shapes)[0];
                Assert.Equal(OdfTextFitMode.None, originalShape.TextFitMode); string[] before = XmlState(original);
                var destination = OdgDocument.Create(); var existing = AddBox(destination.AddPage("Existing"), "Existing");
                var imported = destination.ImportPage(original, 0, "Imported");
                Assert.Null(existing.TextFitMode); Assert.Equal(before, XmlState(original));
                foreach (var actualPage in new[] { imported, Reopen(destination, flat).Pages[1] }) {
                    var actual = (master ? actualPage.MasterShapes : actualPage.Shapes)[0];
                    Assert.Equal(OdfTextFitMode.None, actual.TextFitMode);
                    AssertPair(actual, defaultShrink ? "false" : null, defaultShrink ? null : "false");
                    Assert.False(actual.WrapText);
                }
            }
        }
    }

    [Theory]
    [InlineData(false, OfficeTextAreaAlignment.FullWidth)]
    [InlineData(false, OfficeTextAreaAlignment.Center)]
    [InlineData(true, OfficeTextAreaAlignment.FullWidth)]
    [InlineData(true, OfficeTextAreaAlignment.Right)]
    public void FixedBoxesAndRectangleLabelsShrinkThroughSharedRenderingAndKeepSavedBounds(bool rectangle, OfficeTextAreaAlignment area) {
        var document = OdgDocument.Create(); var shape = AddBox(document.AddPage(), "Fitted", rectangle, height: 65);
        shape.Paragraphs[0].Text = "FIRST\nSECOND\nEND_MARKER"; shape.Paragraphs[0].FontSize = OdfLength.Points(36);
        shape.TextFitMode = OdfTextFitMode.ShrinkToFit; shape.TextAreaAlignment = area;
        shape.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Bottom;
        shape.TextPadding = Padding(2); shape.WrapText = false;
        foreach (bool flat in new[] { false, true }) {
            var actual = Reopen(document, flat); string[] before = XmlState(actual);
            var result = actual.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
            Assert.True(text.ShrinkToFit); Assert.Equal(area, text.TextAreaAlignment);
            Assert.Equal(OfficeTextVerticalAlignment.Bottom, text.VerticalAlignment); Assert.False(text.WrapText);
            Assert.Equal("FIRST\nSECOND\nEND_MARKER", text.PlainText);
            Assert.Equal(20, text.X); Assert.Equal(30, text.Y); Assert.Equal(240, text.Width); Assert.Equal(65, text.Height);
            Assert.Equal(2, text.Padding.Top); Assert.Equal(36, Assert.Single(text.Paragraphs).Runs[0].FontSize);
            Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":text-fitting", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Approximated);
            Assert.DoesNotContain(result.Report.Mappings, m => m.Status is OdfConversionMappingStatus.Unsupported or OdfConversionMappingStatus.Skipped);
            XElement[] rendered = SvgText(result.Value);
            Assert.Equal("FIRSTSECONDEND_MARKER", string.Concat(rendered.Select(element => element.Value)));
            Assert.All(rendered, element => {
                double size = Number(element, "font-size"); Assert.InRange(size, 6, 35.99);
                Assert.InRange(Number(element, "x"), text.X + text.Padding.Left, text.X + text.Width - text.Padding.Right);
                Assert.InRange(Number(element, "y"), text.Y + text.Padding.Top, text.Y + text.Height - text.Padding.Bottom);
            });
            Assert.Throws<OdfConversionLossException>(() => actual.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnAnyLoss));
            Assert.Equal(before, XmlState(actual));
        }
    }

    [Theory]
    [InlineData(OfficeTextAreaAlignment.Left)]
    [InlineData(OfficeTextAreaAlignment.Center)]
    [InlineData(OfficeTextAreaAlignment.Right)]
    public void ShrinkingRectangleAreasHonorWrappingWhileOrdinaryAreasKeepTheirQualifiedProfile(OfficeTextAreaAlignment area) {
        var document = OdgDocument.Create(); var shape = AddBox(document.AddPage(), "WrappedRectangle", rectangle: true);
        shape.Bounds = new OdfRect(OdfLength.Points(20), OdfLength.Points(30), OdfLength.Points(80), OdfLength.Points(100));
        shape.Paragraphs[0].Text = "Alpha beta gamma delta"; shape.Paragraphs[0].FontSize = OdfLength.Points(18);
        shape.TextAreaAlignment = area; shape.WrapText = true; shape.TextPadding = Padding(0);
        foreach (OdfTextFitMode mode in new[] { OdfTextFitMode.ShrinkToFit, OdfTextFitMode.None }) {
            shape.TextFitMode = mode; string[] before = XmlState(document);
            var result = document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
            bool shrinking = mode == OdfTextFitMode.ShrinkToFit;
            Assert.Equal(shrinking, text.WrapText); Assert.Equal(shrinking, text.ShrinkToFit);
            Assert.Equal(area, text.TextAreaAlignment); Assert.Equal("Alpha beta gamma delta", text.PlainText);
            Assert.Equal(!shrinking, result.Report.Mappings.Any(m => m.Feature.EndsWith(":text-area-wrapping", StringComparison.Ordinal) &&
                m.Status == OdfConversionMappingStatus.Approximated));
            int renderedLines = SvgText(result.Value).Select(element => Number(element, "y")).Distinct().Count();
            if (shrinking) Assert.True(renderedLines > 1);
            else Assert.Equal(1, renderedLines);
            Assert.Equal(before, XmlState(document));
        }
    }

    [Theory]
    [InlineData(OfficeTextAreaAlignment.FullWidth)]
    [InlineData(OfficeTextAreaAlignment.Left)]
    [InlineData(OfficeTextAreaAlignment.Center)]
    [InlineData(OfficeTextAreaAlignment.Right)]
    public void UnwrappedTextWiderThanTheFittingFloorReportsLossRegardlessOfItsAreaAnchor(OfficeTextAreaAlignment area) {
        var document = OdgDocument.Create(); var shape = AddBox(document.AddPage(), "TooWide");
        string body = new string('W', 200) + "END_MARKER";
        shape.Paragraphs[0].Text = body; shape.TextAreaAlignment = area; shape.WrapText = false;
        shape.TextFitMode = OdfTextFitMode.ShrinkToFit;
        string[] before = XmlState(document); var result = document.Pages[0].ToDrawing();
        var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Equal(body, text.PlainText); Assert.True(text.ShrinkToFit); Assert.Equal(area, text.TextAreaAlignment);
        AssertClipped(document, result.Report);
        Assert.Equal(before, XmlState(document));
        // Ordinary intrinsic text can intentionally overhang; fitting promises containment instead.
        shape.TextFitMode = OdfTextFitMode.None; before = XmlState(document);
        var ordinary = document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        Assert.Equal(body, Assert.Single(ordinary.Value.Elements.OfType<OfficeDrawingRichText>()).PlainText);
        Assert.DoesNotContain(ordinary.Report.Mappings, m => m.Feature.EndsWith(":text-clipped", StringComparison.Ordinal));
        Assert.Equal(before, XmlState(document));
    }

    [Theory]
    [InlineData(null, false)]
    [InlineData(OdfTextFitMode.None, true)]
    [InlineData(OdfTextFitMode.ShrinkToFit, false)]
    [InlineData(OdfTextFitMode.ShrinkToFit, true)]
    public void ImpossibleFixedHeightIsAnExplicitLossForOrdinaryAndFittedBoxedText(OdfTextFitMode? mode, bool rectangle) {
        var document = OdgDocument.Create(); var shape = AddBox(document.AddPage(), "Clipped", rectangle, height: 4);
        shape.Paragraphs[0].Text = "END_MARKER"; shape.Paragraphs[0].FontSize = OdfLength.Points(12);
        shape.TextFitMode = mode; shape.TextPadding = Padding(0); shape.WrapText = false;
        string[] before = XmlState(document); var result = document.Pages[0].ToDrawing();
        var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Equal("END_MARKER", text.PlainText); Assert.Equal(mode == OdfTextFitMode.ShrinkToFit, text.ShrinkToFit);
        AssertClipped(document, result.Report);
        Assert.DoesNotContain(SvgText(result.Value), element => element.Value.IndexOf("END_MARKER", StringComparison.Ordinal) >= 0);
        Assert.Equal(before, XmlState(document));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OrdinaryEllipseTextReportsClippingThroughTheSameCachedTextFrame(bool explicitNone) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = page.Shapes.AddEllipse(Bounds(height: 4), "ClippedEllipse");
        var paragraph = shape.AddParagraph("END_MARKER");
        paragraph.FontFamily = "Liberation Sans"; paragraph.FontSize = OdfLength.Points(12);
        if (explicitNone) shape.TextFitMode = OdfTextFitMode.None;
        shape.TextPadding = Padding(0); shape.WrapText = false;
        string[] before = XmlState(document); var result = page.ToDrawing();
        var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Equal("END_MARKER", text.PlainText); Assert.False(text.ShrinkToFit);
        AssertClipped(document, result.Report);
        Assert.DoesNotContain(SvgText(result.Value), element => element.Value.IndexOf("END_MARKER", StringComparison.Ordinal) >= 0);
        Assert.Equal(before, XmlState(document));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PaddingAndFixedParagraphMetricsCannotBeRecoveredByFontShrinking(bool consumesFrame) {
        var document = OdgDocument.Create(); var shape = AddBox(document.AddPage(), "Constrained", height: 65);
        shape.TextFitMode = OdfTextFitMode.ShrinkToFit; shape.Paragraphs[0].Text = "END_MARKER";
        if (consumesFrame) shape.TextPadding = Padding(40);
        else shape.Paragraphs[0].EnsureStyle().SetProperty(OdfNamespaces.Style + "paragraph-properties", OdfNamespaces.Fo + "line-height", "80pt");
        string[] before = XmlState(document); var result = document.Pages[0].ToDrawing();
        AssertClipped(document, result.Report);
        if (consumesFrame) Assert.Empty(result.Value.Elements.OfType<OfficeDrawingRichText>());
        else Assert.Equal("END_MARKER", Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).PlainText);
        Assert.Equal(before, XmlState(document));
    }

    [Theory]
    [InlineData("all", null)]
    [InlineData("shrink-to-fit", null)]
    [InlineData("unknown", "false")]
    [InlineData("false", "unknown")]
    [InlineData("true", "true")]
    public void UnknownLegacyAndConflictingPairsPreserveXmlAndRetainFallbackBody(string stretch, string? shrink) {
        var document = OdgDocument.Create(); var shape = AddBox(document.AddPage(), "Unsupported");
        SetPair(GraphicStyle(shape), stretch, shrink);
        foreach (bool flat in new[] { false, true }) {
            var actual = Reopen(document, flat); string[] before = XmlState(actual);
            Assert.Throws<NotSupportedException>(() => { _ = actual.Pages[0].Shapes[0].TextFitMode; });
            var result = actual.Pages[0].ToDrawing();
            AssertUnsupportedFitting(actual, result);
            AssertPair(actual.Pages[0].Shapes[0], stretch, shrink); Assert.Equal(before, XmlState(actual));
        }
    }

    [Theory]
    [InlineData("stretch")]
    [InlineData("auto-grow")]
    [InlineData("ellipse")]
    [InlineData("line")]
    public void HeldFittingProfilesKeepFallbackContentAndRejectStrictProjection(string profile) {
        var document = OdgDocument.Create(); var page = document.AddPage(); OdgShape shape;
        if (profile == "ellipse") {
            shape = page.Shapes.AddEllipse(Bounds(), "Unsupported"); shape.AddParagraph("Body");
        } else if (profile == "line") {
            shape = page.Shapes.AddLine(OdfLength.Points(20), OdfLength.Points(30), OdfLength.Points(260), OdfLength.Points(30), "Unsupported");
            shape.AddParagraph("Body");
        } else shape = AddBox(page, "Unsupported");
        shape.Paragraphs[0].FontFamily = "Liberation Sans"; shape.Paragraphs[0].FontSize = OdfLength.Points(12);
        shape.TextFitMode = profile == "stretch" ? OdfTextFitMode.Stretch : OdfTextFitMode.ShrinkToFit;
        if (profile == "auto-grow") SetGraphic(GraphicStyle(shape), OdfNamespaces.Draw + "auto-grow-height", "true");
        string[] before = XmlState(document); var result = page.ToDrawing();
        AssertUnsupportedFitting(document, result);
        if (profile == "auto-grow") Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":text-auto-size", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Equal(before, XmlState(document));
    }

    private static OdgShape AddBox(OdgPage page, string name, bool rectangle = false, double height = 100) {
        var bounds = Bounds(height); var shape = rectangle ? page.Shapes.AddRectangle(bounds, name) : page.Shapes.AddTextBox(bounds, "Body", name);
        if (rectangle) shape.AddParagraph("Body");
        shape.Paragraphs[0].FontFamily = "Liberation Sans"; shape.Paragraphs[0].FontSize = OdfLength.Points(12);
        return shape;
    }
    private static OdfRect Bounds(double height = 100) => new OdfRect(OdfLength.Points(20), OdfLength.Points(30), OdfLength.Points(240), OdfLength.Points(height));
    private static OdfInsets Padding(double points) => new OdfInsets(OdfLength.Points(points), OdfLength.Points(points), OdfLength.Points(points), OdfLength.Points(points));
    private static OdfStyle GraphicStyle(OdfShape shape) => shape.Document.Styles.FindInPart(OdfStyleFamily.Graphic,
        (string)shape.Element.Attribute(OdfNamespaces.Draw + "style-name")!, shape.PartPath)!;
    private static void SetGraphic(OdfStyle style, XName name, string? value) => style.SetProperty(OdfNamespaces.Style + "graphic-properties", name, value);
    private static void SetPair(OdfStyle style, string? stretch, string? shrink) {
        SetGraphic(style, OdfNamespaces.Draw + "fit-to-size", stretch); SetGraphic(style, OdfNamespaces.Style + "shrink-to-fit", shrink);
    }
    private static void Bind(OdfShape shape, OdfStyle style) {
        shape.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", style.Name); shape.Document.MarkPartDirty(shape.PartPath);
    }
    private static void AssertMode(OdfShape shape, OdfTextFitMode mode, string stretch, string shrink) {
        Assert.Equal(mode, shape.TextFitMode); AssertPair(shape, stretch, shrink);
    }
    private static void AssertPair(OdfShape shape, string? stretch, string? shrink) {
        XElement? properties = GraphicStyle(shape).Element.Element(OdfNamespaces.Style + "graphic-properties");
        Assert.Equal(stretch, (string?)properties?.Attribute(OdfNamespaces.Draw + "fit-to-size"));
        Assert.Equal(shrink, (string?)properties?.Attribute(OdfNamespaces.Style + "shrink-to-fit"));
    }
    private static void AssertClipped(OdgDocument document, OdfConversionReport report) {
        Assert.Contains(report.Mappings, m => m.Feature.EndsWith(":text-clipped", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }
    private static void AssertUnsupportedFitting(OdgDocument document, OdfConversionResult<OfficeDrawing> result) {
        var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Equal("Body", text.PlainText); Assert.False(text.ShrinkToFit);
        Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":text-fitting", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }
    private static XElement[] SvgText(OfficeDrawing drawing) => XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(drawing))
        .Descendants(XNamespace.Get("http://www.w3.org/2000/svg") + "text").ToArray();
    private static double Number(XElement element, string attribute) => double.Parse((string)element.Attribute(attribute)!, CultureInfo.InvariantCulture);
    private static string[] XmlState(OdfDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString() };
    private static OdgDocument Reopen(OdgDocument document, bool flat) {
        if (!flat) return OdgDocument.Load(new MemoryStream(document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource })));
        using var stream = new MemoryStream(); document.SaveFlatXml(stream); stream.Position = 0; return OdgDocument.LoadFlatXml(stream);
    }
}
