using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentShapeTextFrameTests {
    [Theory]
    [InlineData(OfficeTextAreaAlignment.FullWidth, "justify", OdfTextAreaVerticalAlignment.Top, "top", true)]
    [InlineData(OfficeTextAreaAlignment.Left, "left", OdfTextAreaVerticalAlignment.Middle, "middle", false)]
    [InlineData(OfficeTextAreaAlignment.Center, "center", OdfTextAreaVerticalAlignment.Bottom, "bottom", true)]
    [InlineData(OfficeTextAreaAlignment.Right, "right", OdfTextAreaVerticalAlignment.Top, "top", false)]
    public void TypedSettingsPersistAndReachSharedDrawingWithoutChangingSource(
        OfficeTextAreaAlignment area, string nativeArea, OdfTextAreaVerticalAlignment vertical, string nativeVertical, bool wrap) {
        var document = OdgDocument.Create(); var shape = AddTextBox(document.AddPage(), "Frame");
        shape.TextAreaAlignment = area; shape.TextVerticalAlignment = vertical; shape.WrapText = wrap;
        shape.TextPadding = Insets("2pt", "3pt", "5pt", "7pt");
        foreach (bool flat in new[] { false, true }) {
            var reopened = Reopen(document, flat); var actual = reopened.Pages[0].Shapes[0];
            Assert.Equal(area, actual.TextAreaAlignment); Assert.Equal(vertical, actual.TextVerticalAlignment);
            Assert.Equal(wrap, actual.WrapText); Assert.Equal(Insets("2pt", "3pt", "5pt", "7pt"), actual.TextPadding);
            XElement properties = GraphicStyle(actual).Element.Element(OdfNamespaces.Style + "graphic-properties")!;
            Assert.Equal(nativeArea, (string?)properties.Attribute(OdfNamespaces.Draw + "textarea-horizontal-align"));
            Assert.Equal(nativeVertical, (string?)properties.Attribute(OdfNamespaces.Draw + "textarea-vertical-align"));
            Assert.Equal(wrap ? "wrap" : "no-wrap", (string?)properties.Attribute(OdfNamespaces.Fo + "wrap-option"));
            Assert.Equal("2pt", (string?)properties.Attribute(OdfNamespaces.Fo + "padding-top"));
            Assert.Equal("3pt", (string?)properties.Attribute(OdfNamespaces.Fo + "padding-right"));
            Assert.Equal("5pt", (string?)properties.Attribute(OdfNamespaces.Fo + "padding-bottom"));
            Assert.Equal("7pt", (string?)properties.Attribute(OdfNamespaces.Fo + "padding-left"));
            Assert.Null(properties.Attribute(OdfNamespaces.Fo + "padding"));

            string[] before = XmlState(reopened);
            var drawing = reopened.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var text = Assert.Single(drawing.Value.Elements.OfType<OfficeDrawingRichText>());
            Assert.Equal(area, text.TextAreaAlignment); Assert.Equal(wrap, text.WrapText);
            Assert.Equal(vertical switch {
                OdfTextAreaVerticalAlignment.Middle => OfficeTextVerticalAlignment.Center,
                OdfTextAreaVerticalAlignment.Bottom => OfficeTextVerticalAlignment.Bottom,
                _ => OfficeTextVerticalAlignment.Top
            }, text.VerticalAlignment);
            Assert.Equal(7, text.Padding.Left); Assert.Equal(2, text.Padding.Top);
            Assert.Equal(3, text.Padding.Right); Assert.Equal(5, text.Padding.Bottom);
            Assert.Equal("Native\nframe", text.PlainText); Assert.Equal(before, XmlState(reopened));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SharedStylesResolvePaddingPerLevelAndNullRestoresInheritance(bool flat) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var first = AddTextBox(page, "First"); var second = AddTextBox(page, "Second");
        var ancestor = document.Styles.CreateNamed("Ancestor", OdfStyleFamily.Graphic);
        SetGraphic(ancestor, OdfNamespaces.Draw + "textarea-horizontal-align", "left");
        SetGraphic(ancestor, OdfNamespaces.Draw + "textarea-vertical-align", "bottom");
        SetGraphic(ancestor, OdfNamespaces.Fo + "wrap-option", "no-wrap");
        SetGraphic(ancestor, OdfNamespaces.Fo + "padding", "2pt");
        SetGraphic(ancestor, OdfNamespaces.Fo + "padding-left", "7pt");
        var parent = document.Styles.CreateNamed("Parent", OdfStyleFamily.Graphic, ancestor.Name);
        SetGraphic(parent, OdfNamespaces.Fo + "padding", "3pt");
        SetGraphic(parent, OdfNamespaces.Fo + "padding-top", "5pt");
        var shared = document.Styles.CreateAutomatic(OdfStyleFamily.Graphic, parentStyleName: parent.Name);
        SetGraphic(shared, OdfNamespaces.Fo + "padding-right", "11pt");
        SetGraphic(shared, OdfNamespaces.Draw + "textarea-horizontal-align", "center");
        SetGraphic(shared, OdfNamespaces.Draw + "textarea-vertical-align", "middle");
        SetGraphic(shared, OdfNamespaces.Fo + "wrap-option", "wrap");
        first.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", shared.Name);
        second.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", shared.Name);
        document.MarkPartDirty("content.xml");
        document = Reopen(document, flat); first = document.Pages[0].Shapes[0]; second = document.Pages[0].Shapes[1];
        string[] beforeReads = XmlState(document);
        Assert.Equal(Insets("5pt", "11pt", "3pt", "3pt"), first.TextPadding);
        Assert.Equal(OfficeTextAreaAlignment.Center, first.TextAreaAlignment);
        Assert.Equal(OdfTextAreaVerticalAlignment.Middle, first.TextVerticalAlignment); Assert.True(first.WrapText);
        Assert.Equal(beforeReads, XmlState(document));

        first.TextPadding = Insets("13pt", "17pt", "19pt", "23pt");
        first.TextAreaAlignment = OfficeTextAreaAlignment.Right;
        first.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Top; first.WrapText = false;
        Assert.NotEqual(GraphicStyle(first).Name, GraphicStyle(second).Name);
        Assert.Equal(Insets("5pt", "11pt", "3pt", "3pt"), second.TextPadding);
        Assert.Equal(OfficeTextAreaAlignment.Center, second.TextAreaAlignment);
        Assert.Equal(OdfTextAreaVerticalAlignment.Middle, second.TextVerticalAlignment); Assert.True(second.WrapText);
        first.TextPadding = null; first.TextAreaAlignment = null; first.TextVerticalAlignment = null; first.WrapText = null;
        var reopened = Reopen(document, flat); var actual = reopened.Pages[0].Shapes[0];
        Assert.Equal(Insets("5pt", "3pt", "3pt", "3pt"), actual.TextPadding);
        Assert.Equal(OfficeTextAreaAlignment.Left, actual.TextAreaAlignment);
        Assert.Equal(OdfTextAreaVerticalAlignment.Bottom, actual.TextVerticalAlignment); Assert.False(actual.WrapText);
        XElement properties = GraphicStyle(actual).Element.Element(OdfNamespaces.Style + "graphic-properties")!;
        Assert.DoesNotContain(properties.Attributes(), attribute => attribute.Name.Namespace == OdfNamespaces.Fo &&
            attribute.Name.LocalName.StartsWith("padding", StringComparison.Ordinal));
    }

    [Fact]
    public void MissingDeclarationsRemainNullableAndPartialPaddingUsesZeroForMissingSides() {
        var document = OdgDocument.Create(); var shape = AddTextBox(document.AddPage(), "Unset");
        string[] before = XmlState(document);
        Assert.Null(shape.TextAreaAlignment); Assert.Null(shape.TextVerticalAlignment);
        Assert.Null(shape.WrapText); Assert.Null(shape.TextPadding); Assert.Equal(before, XmlState(document));
        SetGraphic(GraphicStyle(shape), OdfNamespaces.Fo + "padding-left", "01.250pt");
        Assert.Equal(Insets("0cm", "0cm", "0cm", "01.250pt"), shape.TextPadding);
        shape.TextPadding = null; Assert.Null(shape.TextPadding);
        var defaults = new XElement(OdfNamespaces.Style + "default-style", new XAttribute(OdfNamespaces.Style + "family", "graphic"),
            new XElement(OdfNamespaces.Style + "graphic-properties", new XAttribute(OdfNamespaces.Fo + "padding", ".50cm"),
                new XAttribute(OdfNamespaces.Draw + "textarea-horizontal-align", "justify"),
                new XAttribute(OdfNamespaces.Draw + "textarea-vertical-align", "top"), new XAttribute(OdfNamespaces.Fo + "wrap-option", "wrap")));
        document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(defaults);
        document.MarkPartDirty("styles.xml");
        before = XmlState(document);
        Assert.Equal(OfficeTextAreaAlignment.FullWidth, shape.TextAreaAlignment);
        Assert.Equal(OdfTextAreaVerticalAlignment.Top, shape.TextVerticalAlignment); Assert.True(shape.WrapText);
        Assert.Equal(Insets(".50cm", ".50cm", ".50cm", ".50cm"), shape.TextPadding); Assert.Equal(before, XmlState(document));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitPaddingReplacesShorthandAndPreservesSchemaUnitsIncludingPixels(bool flat) {
        var document = OdgDocument.Create(); var shape = AddTextBox(document.AddPage(), "Pixels");
        SetGraphic(GraphicStyle(shape), OdfNamespaces.Fo + "padding", "2pt");
        var padding = Insets(".5px", "1.25pc", "2.50mm", "0in");
        shape.TextPadding = padding; shape.TextAreaAlignment = OfficeTextAreaAlignment.FullWidth;
        shape.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Top; shape.WrapText = true;
        var reopened = Reopen(document, flat); var actual = reopened.Pages[0].Shapes[0];
        Assert.Equal(padding, actual.TextPadding);
        Assert.Null(GraphicStyle(actual).Element.Element(OdfNamespaces.Style + "graphic-properties")!.Attribute(OdfNamespaces.Fo + "padding"));
        string[] before = XmlState(reopened); var result = reopened.Pages[0].ToDrawing();
        Assert.Empty(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Status == OdfConversionMappingStatus.Skipped &&
            mapping.Message != null && mapping.Message.Contains("padding", StringComparison.Ordinal));
        Assert.Throws<OdfConversionLossException>(() => reopened.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal(before, XmlState(reopened));
        actual.TextPadding = null; Assert.Null(actual.TextPadding);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MasterEditsStayInTheirPartAndCloneImportCopiesRemainIndependent(bool flat) {
        var source = OdgDocument.Create(); var page = source.AddPage("Source");
        var local = AddTextBox(page, "Local"); var master = page.MasterShapes.AddTextBox(Bounds(), "Master", "Master");
        Configure(local, OfficeTextAreaAlignment.Center); Configure(master, OfficeTextAreaAlignment.Left);
        // Automatic style names are scoped by XML part, including imported collisions.
        GraphicStyle(master).Element.SetAttributeValue(OdfNamespaces.Style + "name", "SameName");
        master.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", "SameName");
        GraphicStyle(local).Element.SetAttributeValue(OdfNamespaces.Style + "name", "SameName");
        local.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", "SameName");
        source.MarkPartDirty("content.xml"); source.MarkPartDirty("styles.xml");
        var clone = source.ClonePage(0, "Copy"); clone.MasterPageName = source.CloneMasterPage(page.MasterPageName, "Independent");
        source = Reopen(source, flat);
        byte[] contentBefore = source.GetPackageEntryBytes("content.xml");
        source.Pages[0].MasterShapes[0].TextAreaAlignment = OfficeTextAreaAlignment.Right;
        source.Pages[0].MasterShapes[0].TextPadding = Insets("2pt", "3pt", "5pt", "7pt");
        Assert.Equal(contentBefore, source.GetPackageEntryBytes("content.xml"));
        var actual = Reopen(source, flat);
        Assert.Equal(OfficeTextAreaAlignment.Right, actual.Pages[0].MasterShapes[0].TextAreaAlignment);
        Assert.Equal(OfficeTextAreaAlignment.Left, actual.Pages[1].MasterShapes[0].TextAreaAlignment);
        Assert.Equal(OfficeTextAreaAlignment.Center, actual.Pages[0].Shapes[0].TextAreaAlignment);
        actual.Pages[1].Shapes[0].WrapText = false;
        Assert.True(actual.Pages[0].Shapes[0].WrapText);

        var destination = OdgDocument.Create(); destination.AddPage("Existing");
        string[] sourceBefore = XmlState(actual); var imported = destination.ImportPage(actual, 0, "Imported");
        Assert.Equal(sourceBefore, XmlState(actual));
        Assert.Equal(OfficeTextAreaAlignment.Center, imported.Shapes[0].TextAreaAlignment);
        Assert.Equal(OfficeTextAreaAlignment.Right, imported.MasterShapes[0].TextAreaAlignment);
        Assert.Equal(Insets("2pt", "3pt", "5pt", "7pt"), imported.MasterShapes[0].TextPadding);
        imported.Shapes[0].TextAreaAlignment = OfficeTextAreaAlignment.FullWidth;
        imported.MasterShapes[0].TextVerticalAlignment = OdfTextAreaVerticalAlignment.Bottom;
        var destinationAgain = Reopen(destination, flat);
        Assert.Equal(OfficeTextAreaAlignment.FullWidth, destinationAgain.Pages[1].Shapes[0].TextAreaAlignment);
        Assert.Equal(OdfTextAreaVerticalAlignment.Bottom, destinationAgain.Pages[1].MasterShapes[0].TextVerticalAlignment);
        Assert.Equal(sourceBefore, XmlState(actual));
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    public void ImportPreservesPaddingShorthandPrecedenceOverDefaultsAndUnrelatedInheritedSides(bool flat, bool master, bool named) {
        var source = OdgDocument.Create(); var page = source.AddPage("Source");
        var shapes = master ? page.MasterShapes : page.Shapes;
        var shorthand = shapes.AddTextBox(Bounds(), "Shorthand", "Shorthand");
        var side = shapes.AddTextBox(Bounds(), "Side", "Side");
        foreach (OdgShape shape in new[] { shorthand, side }) {
            OdfStyle style;
            if (named) {
                style = source.Styles.CreateNamed(shape.Name + "Style", OdfStyleFamily.Graphic);
                shape.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", style.Name);
                source.MarkPartDirty(shape.PartPath);
            } else style = GraphicStyle(shape);
            SetGraphic(style, OdfNamespaces.Fo + (shape == shorthand ? "padding" : "padding-right"), "3pt");
        }
        source.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(
            new XElement(OdfNamespaces.Style + "default-style", new XAttribute(OdfNamespaces.Style + "family", "graphic"),
                new XElement(OdfNamespaces.Style + "graphic-properties", new XAttribute(OdfNamespaces.Fo + "padding", "5pt"),
                    new XAttribute(OdfNamespaces.Fo + "padding-left", "9pt"), new XAttribute(OdfNamespaces.Fo + "wrap-option", "no-wrap"),
                    new XAttribute(OdfNamespaces.Draw + "textarea-vertical-align", "bottom"))));
        source.MarkPartDirty("styles.xml"); source = Reopen(source, flat);
        var originalShapes = master ? source.Pages[0].MasterShapes : source.Pages[0].Shapes;
        Assert.Equal(Insets("3pt", "3pt", "3pt", "3pt"), originalShapes[0].TextPadding);
        Assert.Equal(Insets("5pt", "3pt", "5pt", "9pt"), originalShapes[1].TextPadding);
        string[] sourceBefore = XmlState(source);
        var destination = OdgDocument.Create(); var existing = AddTextBox(destination.AddPage("Existing"), "Existing");
        var imported = destination.ImportPage(source, 0, "Imported");
        Assert.Equal(sourceBefore, XmlState(source)); Assert.Null(existing.TextPadding); Assert.Null(existing.WrapText);
        foreach (var actualPage in new[] { imported, Reopen(destination, flat).Pages[1] }) {
            var actualShapes = master ? actualPage.MasterShapes : actualPage.Shapes;
            Assert.Equal(Insets("3pt", "3pt", "3pt", "3pt"), actualShapes[0].TextPadding);
            Assert.Equal(Insets("5pt", "3pt", "5pt", "9pt"), actualShapes[1].TextPadding);
            Assert.All(actualShapes, shape => {
                Assert.False(shape.WrapText); Assert.Equal(OdfTextAreaVerticalAlignment.Bottom, shape.TextVerticalAlignment);
            });
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ImportPreservesRenderedParagraphMarginShorthandsAndInheritedExplicitSides(bool flat, bool named) {
        var source = OdgDocument.Create(); var page = source.AddPage("Source");
        var shorthand = AddTextBox(page, "Shorthand"); var side = AddTextBox(page, "Side");
        foreach (OdgShape shape in new[] { shorthand, side }) {
            var paragraph = shape.Paragraphs[0];
            var style = named ? source.Styles.CreateNamed(shape.Name + "Paragraph", OdfStyleFamily.Paragraph) : paragraph.EnsureStyle();
            if (named) {
                style.FontFamily = "Liberation Sans"; style.FontSize = OdfLength.Points(12); paragraph.StyleName = style.Name;
            }
            style.SetProperty(OdfNamespaces.Style + "paragraph-properties",
                OdfNamespaces.Fo + (shape == shorthand ? "margin" : "margin-right"), "3pt");
        }
        source.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(
            new XElement(OdfNamespaces.Style + "default-style", new XAttribute(OdfNamespaces.Style + "family", "paragraph"),
                new XElement(OdfNamespaces.Style + "paragraph-properties", new XAttribute(OdfNamespaces.Fo + "margin", "5pt"),
                    new XAttribute(OdfNamespaces.Fo + "margin-left", "9pt"))));
        source.MarkPartDirty("styles.xml"); source = Reopen(source, flat);
        string[] sourceBefore = XmlState(source); AssertMargins(source.Pages[0]); Assert.Equal(sourceBefore, XmlState(source));
        var destination = OdgDocument.Create(); var imported = destination.ImportPage(source, 0, "Imported");
        Assert.Equal(sourceBefore, XmlState(source));
        AssertMargins(imported); AssertMargins(Reopen(destination, flat).Pages[0]);

        static void AssertMargins(OdgPage actualPage) {
            var texts = actualPage.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value.Elements.OfType<OfficeDrawingRichText>().ToArray();
            Assert.Equal(2, texts.Length);
            // A default shorthand captured above a named parent must not hide that parent's explicit right side.
            var sideMargins = Assert.Single(texts[1].Paragraphs).Margins;
            Assert.Equal(9, sideMargins.Left); Assert.Equal(5, sideMargins.Top);
            Assert.Equal(3, sideMargins.Right); Assert.Equal(5, sideMargins.Bottom);
            var shorthandMargins = Assert.Single(texts[0].Paragraphs).Margins;
            Assert.Equal(3, shorthandMargins.Left); Assert.Equal(3, shorthandMargins.Top);
            Assert.Equal(3, shorthandMargins.Right); Assert.Equal(3, shorthandMargins.Bottom);
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ImportPreservesSavedNativeBorderAndLineWidthCascade(bool flat, bool paragraph) {
        var source = OdgDocument.Create(); var page = source.AddPage("Source");
        var shorthand = AddTextBox(page, "Shorthand"); var side = AddTextBox(page, "Side");
        XName propertiesName = OdfNamespaces.Style + (paragraph ? "paragraph-properties" : "graphic-properties");
        foreach (OdgShape shape in new[] { shorthand, side }) {
            var style = paragraph ? source.Styles.CreateNamed(shape.Name + "Paragraph", OdfStyleFamily.Paragraph) : GraphicStyle(shape);
            if (paragraph) shape.Paragraphs[0].StyleName = style.Name;
            bool allSides = shape == shorthand;
            style.SetProperty(propertiesName, OdfNamespaces.Fo + (allSides ? "border" : "border-right"),
                allSides ? "6pt double #112233" : "9pt double #334455");
            style.SetProperty(propertiesName, OdfNamespaces.Style + (allSides ? "border-line-width" : "border-line-width-right"),
                allSides ? "1pt 2pt 3pt" : "2pt 3pt 4pt");
        }
        source.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(
            new XElement(OdfNamespaces.Style + "default-style", new XAttribute(OdfNamespaces.Style + "family", paragraph ? "paragraph" : "graphic"),
                new XElement(propertiesName, new XAttribute(OdfNamespaces.Fo + "border", "18pt double #556677"),
                    new XAttribute(OdfNamespaces.Fo + "border-left", "30pt double #778899"),
                    new XAttribute(OdfNamespaces.Style + "border-line-width", "5pt 6pt 7pt"),
                    new XAttribute(OdfNamespaces.Style + "border-line-width-left", "9pt 10pt 11pt"))));
        source.MarkPartDirty("styles.xml"); source = Reopen(source, flat);
        string[] sourceBefore = XmlState(source); AssertEdges(source.Pages[0]); Assert.Equal(sourceBefore, XmlState(source));
        var destination = OdgDocument.Create(); var imported = destination.ImportPage(source, 0, "Imported");
        Assert.Equal(sourceBefore, XmlState(source));
        AssertEdges(imported); AssertEdges(Reopen(destination, flat).Pages[0]);

        void AssertEdges(OdgPage actualPage) {
            // These are native saved style semantics; shared Draw rendering does not claim paragraph/frame border support.
            Assert.Equal(new[] { "18pt double #556677", "9pt double #334455", "18pt double #556677", "30pt double #778899" },
                NativeEdges(actualPage.Shapes[1], OdfNamespaces.Fo + "border"));
            Assert.Equal(new[] { "5pt 6pt 7pt", "2pt 3pt 4pt", "5pt 6pt 7pt", "9pt 10pt 11pt" },
                NativeEdges(actualPage.Shapes[1], OdfNamespaces.Style + "border-line-width"));
            Assert.Equal(Enumerable.Repeat("6pt double #112233", 4), NativeEdges(actualPage.Shapes[0], OdfNamespaces.Fo + "border"));
            Assert.Equal(Enumerable.Repeat("1pt 2pt 3pt", 4), NativeEdges(actualPage.Shapes[0], OdfNamespaces.Style + "border-line-width"));
        }
        string?[] NativeEdges(OdgShape shape, XName common) => new[] { "top", "right", "bottom", "left" }.Select(edge => {
            XName name = common.Namespace + (common.LocalName + "-" + edge);
            return paragraph ? shape.Paragraphs[0].Styles.Select(style =>
                (string?)style.Element.Element(propertiesName)?.Attribute(name) ?? (string?)style.Element.Element(propertiesName)?.Attribute(common))
                .FirstOrDefault(value => value != null) : shape.ReadGraphicProperty(name, common);
        }).ToArray();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PresentationTextBoxUsesTheSameNativeTextFrameProperties(bool flat) {
        var presentation = OdpPresentation.Create(); var shape = presentation.AddSlide("Text").AddTextBox(Bounds(), "Presentation text");
        Configure(shape, OfficeTextAreaAlignment.Right); shape.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Justify;
        OdpPresentation reopened;
        if (flat) {
            using var stream = new MemoryStream(); presentation.SaveFlatXml(stream); stream.Position = 0;
            reopened = OdpPresentation.LoadFlatXml(stream);
        } else reopened = OdpPresentation.Load(new MemoryStream(presentation.ToBytes()));
        var actual = Assert.IsType<OdpTextBox>(Assert.Single(reopened.Slides[0].Shapes)); string[] before = XmlState(reopened);
        Assert.Equal(OfficeTextAreaAlignment.Right, actual.TextAreaAlignment);
        Assert.Equal(OdfTextAreaVerticalAlignment.Justify, actual.TextVerticalAlignment); Assert.True(actual.WrapText);
        Assert.Equal(Insets("1pt", "2pt", "3pt", "4pt"), actual.TextPadding);
        Assert.Equal("Presentation text", Assert.Single(actual.Paragraphs).Text); Assert.Equal(before, XmlState(reopened));
    }

    [Theory]
    [InlineData("horizontal")]
    [InlineData("vertical")]
    [InlineData("wrap")]
    public void UnknownImportedDeclarationsThrowWithoutChangingEitherXmlPart(string property) {
        var document = OdgDocument.Create(); var shape = AddTextBox(document.AddPage(), "Unknown");
        XName name = property == "horizontal" ? OdfNamespaces.Draw + "textarea-horizontal-align" :
            property == "vertical" ? OdfNamespaces.Draw + "textarea-vertical-align" : OdfNamespaces.Fo + "wrap-option";
        SetGraphic(GraphicStyle(shape), name, "unknown");
        foreach (bool flat in new[] { false, true }) {
            var reopened = Reopen(document, flat); var actual = reopened.Pages[0].Shapes[0]; string[] before = XmlState(reopened);
            Assert.Throws<NotSupportedException>(() => {
                if (property == "horizontal") _ = actual.TextAreaAlignment;
                else if (property == "vertical") _ = actual.TextVerticalAlignment;
                else _ = actual.WrapText;
            });
            Assert.Equal("unknown", (string?)GraphicStyle(actual).Element.Element(OdfNamespaces.Style + "graphic-properties")!.Attribute(name));
            Assert.Equal(before, XmlState(reopened));
        }
    }

    [Fact]
    public void InvalidSettersRejectBeforeChangingTheShapeReferenceOrSharedStyle() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = AddTextBox(page, "Edited"); var sibling = AddTextBox(page, "Sibling");
        sibling.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", GraphicStyle(shape).Name);
        document.MarkPartDirty("content.xml"); string[] before = XmlState(document);
        Assert.Throws<ArgumentOutOfRangeException>(() => shape.TextAreaAlignment = (OfficeTextAreaAlignment)99);
        Assert.Throws<ArgumentOutOfRangeException>(() => shape.TextVerticalAlignment = (OdfTextAreaVerticalAlignment)99);
        foreach (string invalid in new[] { "-1pt", "+1pt", "1e2pt", "1PT", "1%", "1em", "NaNpt", "Infinitypt", "1pt extra",
            new string('9', 400) + "px", new string('9', 308) + "in" }) {
            // Exercise each setter edge; a late invalid edge must not commit earlier valid edges.
            for (int side = 0; side < 4; side++) {
                var lengths = new[] { OdfLength.Points(1), OdfLength.Points(2), OdfLength.Points(3), OdfLength.Points(4) };
                lengths[side] = OdfLength.Parse(invalid);
                Assert.Throws<ArgumentException>(() => shape.TextPadding = new OdfInsets(lengths[0], lengths[1], lengths[2], lengths[3]));
                Assert.Equal(before, XmlState(document));
            }
        }
        Assert.Equal(before, XmlState(document));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeVerticalJustificationIsPreservedWithExplicitProjectionLoss(bool flat) {
        var document = OdgDocument.Create(); var shape = AddTextBox(document.AddPage(), "Justified");
        Configure(shape, OfficeTextAreaAlignment.FullWidth); shape.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Justify;
        var reopened = Reopen(document, flat); var actual = reopened.Pages[0].Shapes[0];
        Assert.Equal(OdfTextAreaVerticalAlignment.Justify, actual.TextVerticalAlignment);
        Assert.Equal("justify", (string?)GraphicStyle(actual).Element.Element(OdfNamespaces.Style + "graphic-properties")!
            .Attribute(OdfNamespaces.Draw + "textarea-vertical-align"));
        string[] before = XmlState(reopened); var result = reopened.Pages[0].ToDrawing();
        Assert.Empty(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Status == OdfConversionMappingStatus.Skipped &&
            mapping.Message != null && mapping.Message.Contains("vertical alignment", StringComparison.Ordinal));
        Assert.Throws<OdfConversionLossException>(() => reopened.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal(before, XmlState(reopened));
    }

    private static OdgShape AddTextBox(OdgPage page, string name) {
        var shape = page.Shapes.AddTextBox(Bounds(), "", name); var paragraph = shape.Paragraphs[0];
        paragraph.Text = "Native\nframe"; paragraph.FontFamily = "Liberation Sans"; paragraph.FontSize = OdfLength.Points(12);
        return shape;
    }
    private static void Configure(OdfShape shape, OfficeTextAreaAlignment area) {
        shape.TextAreaAlignment = area; shape.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Middle;
        shape.WrapText = true; shape.TextPadding = Insets("1pt", "2pt", "3pt", "4pt");
    }
    private static OdfRect Bounds() => new OdfRect(OdfLength.Points(20), OdfLength.Points(30), OdfLength.Points(240), OdfLength.Points(100));
    private static OdfInsets Insets(string top, string right, string bottom, string left) =>
        new OdfInsets(OdfLength.Parse(top), OdfLength.Parse(right), OdfLength.Parse(bottom), OdfLength.Parse(left));
    private static OdfStyle GraphicStyle(OdfShape shape) => shape.Document.Styles.FindInPart(OdfStyleFamily.Graphic,
        (string)shape.Element.Attribute(OdfNamespaces.Draw + "style-name")!, shape.PartPath)!;
    private static void SetGraphic(OdfStyle style, XName name, string value) => style.SetProperty(OdfNamespaces.Style + "graphic-properties", name, value);
    private static string[] XmlState(OdfDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString() };
    private static OdgDocument Reopen(OdgDocument document, bool flat) {
        if (!flat) return OdgDocument.Load(new MemoryStream(document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource })));
        using var stream = new MemoryStream(); document.SaveFlatXml(stream); stream.Position = 0; return OdgDocument.LoadFlatXml(stream);
    }
}
