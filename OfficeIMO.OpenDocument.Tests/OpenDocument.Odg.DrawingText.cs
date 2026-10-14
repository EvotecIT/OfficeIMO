using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgDrawingTextTests {
    [Fact]
    public void ProjectsMixedRunsParagraphsAndSafeLinksWithoutMutatingNativeXml() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8), "Rich");
        shape.FontFamily = "Liberation Sans"; shape.FontSize = OdfLength.Points(16);
        var first = shape.AddParagraph("Regular "); first.TextAlign = "center"; first.MarginLeft = OdfLength.Points(8); first.LineHeight = OdfLength.Parse("150%");
        var parent = first.AddRun("Bold "); parent.Bold = true; parent.Color = OdfColor.Parse("#123456");
        var nested = parent.AddRun("Small"); nested.FontSize = OdfLength.Parse("75%"); nested.Italic = true;
        var link = first.AddHyperlink(" Help", "https://example.invalid/help"); link.Underline = true;
        var last = shape.AddParagraph("Last"); last.TextAlign = "end"; last.LineHeight = OdfLength.Points(22); last.MarginTop = OdfLength.Points(3);
        string xml = document.GetXml("content.xml").ToString();
        var result = page.ToDrawing(); var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Equal(xml, document.GetXml("content.xml").ToString());
        Assert.Equal("Regular Bold Small Help\nLast", text.PlainText); Assert.Equal(2, text.Paragraphs.Count);
        Assert.Equal(OfficeTextAlignment.Center, text.Paragraphs[0].Alignment); Assert.Equal(1.5, text.Paragraphs[0].LineHeightFactor);
        Assert.Equal(8, text.Paragraphs[0].Margins.Left); Assert.Equal(22, text.Paragraphs[1].LineHeight);
        var small = text.Paragraphs[0].Runs.Single(r => r.Text == "Small");
        Assert.Equal(12, small.FontSize); Assert.True(small.Bold); Assert.True(small.Italic); Assert.Equal(OfficeColor.Parse("#123456"), small.Color);
        Assert.Equal("Liberation Sans", small.FontFamily);
        var linked = text.Paragraphs[0].Runs.Single(r => r.Text == "Help"); Assert.True(linked.Underline);
        Assert.False(result.Report.HasSkippedOrUnsupported);
        string svg = OfficeDrawingSvgExporter.ToSvg(result.Value);
        Assert.Contains("https://example.invalid/help", svg); Assert.Contains("#123456", svg);
    }

    [Fact]
    public void NearShorthandAndExplicitDisabledEffectsOverrideInheritedProperties() {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8));
        var baseStyle = document.Styles.CreateNamed("Base", OdfStyleFamily.Paragraph);
        baseStyle.MarginLeft = OdfLength.Points(40); baseStyle.LineHeight = OdfLength.Points(40);
        baseStyle.SetProperty(OdfNamespaces.Style + "text-properties", OdfNamespaces.Fo + "text-shadow", "1pt 1pt");
        var p = shape.AddParagraph("Override"); p.StyleName = baseStyle.Name;
        var near = p.EnsureStyle(); near.SetProperty(OdfNamespaces.Style + "paragraph-properties", OdfNamespaces.Fo + "margin", "3pt");
        near.SetProperty(OdfNamespaces.Style + "paragraph-properties", OdfNamespaces.Fo + "line-height", "normal");
        near.SetProperty(OdfNamespaces.Style + "text-properties", OdfNamespaces.Fo + "text-shadow", "none");
        shape.FontSize = OdfLength.Points(12);
        var graphicBase = document.Styles.Find(OdfStyleFamily.Graphic, (string)shape.Element.Attribute(OdfNamespaces.Draw + "style-name")!)!;
        graphicBase.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "padding-left", "20pt");
        // A distinct child graphic shorthand must supersede the parent's side-specific inset.
        var child = document.Styles.CreateAutomatic(OdfStyleFamily.Graphic); child.ParentStyleName = graphicBase.Name; shape.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", child.Name);
        child.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "padding", "4pt");
        var result = document.Pages[0].ToDrawing(); var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Equal(3, text.Paragraphs[0].Margins.Left); Assert.Null(text.Paragraphs[0].LineHeight); Assert.Equal(4, text.Padding.Left);
        Assert.DoesNotContain(result.Report.Mappings, m => m.Feature.EndsWith(":text-shadow", StringComparison.Ordinal));
    }

    [Fact]
    public void ListsFieldsUnsafeLinksAndTypographyLossesAreExplicitAndStrictPolicyRejectsThem() {
        var document = OdgDocument.Create(); var page = document.AddPage(); var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8), "Limits");
        var p = shape.AddParagraph("Text "); p.AddHyperlink("Unsafe", "javascript:alert(1)"); p.AddField(OdfTextFieldKind.Date, "Cached date");
        p.AddRun("Caps").SmallCaps = true; p.AddText("\tTabs"); shape.AddList(true).AddItem("Item");
        var result = page.ToDrawing(); string svg = OfficeDrawingSvgExporter.ToSvg(result.Value);
        Assert.Contains("Unsafe", svg); Assert.DoesNotContain("javascript:", svg);
        Assert.Contains("1.", svg);
        Assert.Contains(result.Report.Mappings, m => m.Feature == "shape:Limits:text:field-date-time-cache" && m.Status == OdfConversionMappingStatus.Approximated);
        foreach (string feature in new[] { "hyperlink-target", "small-caps" })
            Assert.Contains(result.Report.Mappings, m => m.Feature == "shape:Limits:text:" + feature && m.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal("javascript:alert(1)", p.Hyperlinks.Single().Href);
    }

    [Theory]
    [InlineData("text-underline-mode", "skip-white-space", "underline-mode")]
    [InlineData("text-line-through-mode", "skip-white-space", "line-through-mode")]
    [InlineData("text-underline-color", "#FF0000", "underline-color")]
    [InlineData("text-line-through-color", "#FF0000", "line-through-color")]
    [InlineData("text-underline-width", "bold", "underline-width")]
    [InlineData("text-line-through-width", "bold", "line-through-width")]
    [InlineData("font-size-rel", "+6pt", "relative-font-size")]
    public void UnmappedNativeTypographyCannotPassTheOmissionRejectingPolicy(string property, string value, string feature) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8), "Formatting");
        var parent = document.Styles.CreateNamed("Parent", OdfStyleFamily.Paragraph); parent.FontSize = OdfLength.Points(12);
        var named = document.Styles.CreateNamed("Native", OdfStyleFamily.Paragraph); named.ParentStyleName = parent.Name;
        named.Underline = true; named.StrikeThrough = true;
        named.SetProperty(OdfNamespaces.Style + "text-properties", OdfNamespaces.Style + property, value);
        shape.AddParagraph("One two").StyleName = named.Name;
        var result = page.ToDrawing();
        Assert.Contains(result.Report.Mappings, m => m.Feature == "shape:Formatting:text:" + feature && m.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal("One two", Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).PlainText);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AnEmptyPageNumberCacheResolvesForShapeAndLineLabels(bool lineLabel) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = lineLabel ? page.Shapes.AddLine(OdfLength.Points(10), OdfLength.Points(10), OdfLength.Points(100), OdfLength.Points(10)) :
            page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8));
        shape.Name = "EmptyField";
        shape.AddParagraph().AddField(OdfTextFieldKind.PageNumber, string.Empty);
        var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>());
        Assert.Equal("1", Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).PlainText);
        Assert.Contains(result.Report.Mappings, m => m.Feature == "shape:EmptyField:text:field-page-number" && m.Status == OdfConversionMappingStatus.Approximated);
    }

    [Fact]
    public void ExplicitFontAndDecorationOverridesDisableInheritedTypographyLosses() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8));
        var parent = document.Styles.CreateNamed("Native", OdfStyleFamily.Paragraph); parent.FontSize = OdfLength.Points(12);
        parent.Underline = true; parent.StrikeThrough = true;
        foreach (var declaration in new[] { ("font-size-rel", "+6pt"), ("text-underline-mode", "skip-white-space"),
            ("text-line-through-mode", "skip-white-space"), ("text-underline-color", "#FF0000"), ("text-line-through-width", "bold") })
            parent.SetProperty(OdfNamespaces.Style + "text-properties", OdfNamespaces.Style + declaration.Item1, declaration.Item2);
        var p = shape.AddParagraph("Override"); p.StyleName = parent.Name;
        p.FontSize = OdfLength.Points(9); p.StrikeThrough = false;
        var child = p.EnsureStyle();
        child.SetProperty(OdfNamespaces.Style + "text-properties", OdfNamespaces.Style + "text-underline-mode", "continuous");
        child.SetProperty(OdfNamespaces.Style + "text-properties", OdfNamespaces.Style + "text-underline-color", "font-color");
        var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        var run = Assert.Single(Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).Paragraphs[0].Runs);
        Assert.Equal(9, run.FontSize); Assert.True(run.Underline); Assert.False(run.Strikethrough);
        Assert.False(result.Report.HasSkippedOrUnsupported);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnrenderableTextDoesNotDiscardSupportedShapeGeometry(bool overLimit) {
        var document = OdgDocument.Create(); var page = document.AddPage(); var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 8, 4), "Retained");
        if (overLimit) shape.AddParagraph().Element.Add(new XElement(OdfNamespaces.Text + "s", new XAttribute(OdfNamespaces.Text + "c", "100001")));
        else { shape.AddParagraph("Bad font"); shape.FontSize = OdfLength.Parse("unsupported"); }
        var result = page.ToDrawing(); Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>());
        Assert.Empty(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Contains(result.Report.Mappings, m => m.Feature == "shape:Retained:text" && m.Status == OdfConversionMappingStatus.Skipped);
    }

    [Fact]
    public void IndependentProducerTextOpacityProjectsAlongsideLiteralGeometry() {
        var document = OdgDocument.LoadFlatXml(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-transparent-text.fodg"));
        var result = document.Pages[0].ToDrawing();
        var run = Assert.Single(Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).Paragraphs[0].Runs);
        Assert.Equal("asdf", run.Text); Assert.Equal(66, run.FontSize); Assert.Equal(OfficeColor.FromRgba(255, 0, 0, 64), run.Color);
        Assert.DoesNotContain(result.Report.Mappings, m => m.Feature.EndsWith(":text-opacity", StringComparison.Ordinal));
        Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>());
        Assert.Contains("asdf", OfficeDrawingSvgExporter.ToSvg(result.Value));
    }

    [Fact]
    public void OversizedLineLabelRetainsItsGeometryAndReportsTextSeparately() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var line = page.Shapes.AddLine(OdfLength.Points(10), OdfLength.Points(10), OdfLength.Points(100), OdfLength.Points(100));
        line.Name = "Line"; line.AddParagraph().Element.Add(new XElement(OdfNamespaces.Text + "s", new XAttribute(OdfNamespaces.Text + "c", "100001")));
        var result = page.ToDrawing(); Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>());
        Assert.Contains(result.Report.Mappings, m => m.Feature == "shape:Line:text" && m.Status == OdfConversionMappingStatus.Skipped);
        Assert.DoesNotContain(result.Report.Mappings, m => m.Feature == "shape:Line" && m.Status == OdfConversionMappingStatus.Skipped);
    }
}
