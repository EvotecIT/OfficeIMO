using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using A = DocumentFormat.OpenXml.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class PowerPointOdpReviewWaveTests {
    private static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    private static readonly XNamespace Style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    private static readonly XNamespace Svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
    private static readonly XNamespace Table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";

    [Fact]
    public void WhitespaceSlideNameDoesNotCollideWithGeneratedName() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank).Name = " ";
        source.AddSlide(PowerPointSlideLayoutType.Blank).Name = "Slide1";

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.Equal(new[] { "Slide1", "Slide1_2" }, conversion.Value.Slides.Select(slide => slide.Name));
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "slide-names" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Theory]
    [InlineData("collapse")]
    [InlineData("filter")]
    public void OdpShapeAccessibilityAndHiddenTableRemainExplicitLoss(string visibility) {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide();
        OdpRectangle shape = slide.AddRectangle(OdfRect.FromCentimeters(1, 1, 3, 2));
        shape.Element.Add(new XElement(Svg + "title", "Accessible title"),
            new XElement(Svg + "desc", "Accessible description"));
        OdpTable table = slide.AddTable(OdfRect.FromCentimeters(1, 4, 8, 4), 2, 2);
        table.Element.Descendants(Table + "table-row").First()
            .SetAttributeValue(Table + "visibility", visibility);
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation converted = conversion.Value;
        AssertLoss(conversion.Report, "shape-accessibility");
        AssertLoss(conversion.Report, "table-visibility");
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void EmbeddedOdpFontFaceRemainsExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide();
        source.Package.GetXml("styles.xml").Root!.Element(Office + "font-face-decls")!.Add(
            new XElement(Style + "font-face", new XAttribute(Style + "name", "Embedded"),
                new XElement(Svg + "font-face-src",
                    new XElement(Svg + "font-face-uri"))));
        source.Package.MarkXmlDirty("styles.xml");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation converted = conversion.Value;
        AssertLoss(conversion.Report, "embedded-fonts");
    }

    [Fact]
    public void ThemeBackgroundAndAdvanceOnlyTransitionRemainExplicitLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        Slide slide = source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!;
        slide.CommonSlideData!.Background = new Background(new BackgroundStyleReference(
            new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = 1U });
        slide.Transition = new Transition { AdvanceAfterTime = "5000" };

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        AssertLoss(conversion.Report, "slide-backgrounds");
        AssertLoss(conversion.Report, "slide-transition-timing");
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void SystemColorSlideBackgroundRemainsExplicitLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        Slide slide = source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!;
        slide.CommonSlideData!.Background = new Background(new BackgroundProperties(new A.SolidFill(
            new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "336699" })));

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        AssertLoss(conversion.Report, "slide-backgrounds");
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void AuthoredPowerPointParagraphLayoutRemainsExplicitLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        PowerPointTextBox box = source.AddSlide(PowerPointSlideLayoutType.Blank)
            .AddTextBoxPoints("Indented", 20, 20, 200, 40);
        PowerPointParagraph paragraph = box.Paragraphs.Single();
        paragraph.LeftMarginPoints = 24;
        paragraph.IndentPoints = -12;
        paragraph.SpaceBeforePoints = 8;
        source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<A.ParagraphProperties>().First().DefaultTabSize = 914400;

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        AssertLoss(conversion.Report, "paragraph-layout");
    }

    [Fact]
    public void PowerPointDefaultTabIntervalAloneRemainsExplicitLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank).AddTextBoxPoints("A\tB", 20, 20, 200, 40);
        A.Paragraph paragraph = source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<A.Paragraph>().Single();
        paragraph.ParagraphProperties ??= new A.ParagraphProperties();
        paragraph.ParagraphProperties.DefaultTabSize = 914400;

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        AssertLoss(conversion.Report, "paragraph-layout");
    }

    private static void AssertLoss(OdfConversionReport report, string feature) =>
        Assert.Contains(report.Mappings, mapping => mapping.Feature == feature &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
}
