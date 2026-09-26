using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using A = DocumentFormat.OpenXml.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class PowerPointOdpLossClosureTests {
    [Fact]
    public void DirectEastAsianFontIsReportedBeforeStrictConversion() {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].AddTextBoxPoints("Text", 20, 20, 200, 40);
        A.Run run = source.OpenXmlDocument.PresentationPart!.SlideParts.Single()
            .Slide!.Descendants<A.Run>().Single();
        run.RunProperties = new A.RunProperties(new A.EastAsianFont { Typeface = "Noto Sans CJK JP" });

        AssertPowerPointLoss(source, "text-typography");
    }

    [Fact]
    public void LatinFontMetadataBeyondTypefaceIsReported() {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].AddTextBoxPoints("Text", 20, 20, 200, 40);
        A.Run run = source.OpenXmlDocument.PresentationPart!.SlideParts.Single()
            .Slide!.Descendants<A.Run>().Single();
        run.RunProperties = new A.RunProperties(new A.LatinFont {
            Typeface = "Arial", Panose = "020B0604020202020204"
        });

        AssertPowerPointLoss(source, "text-typography");
    }

    [Fact]
    public void PresentationDefaultTextStyleIsReported() {
        using PowerPointPresentation source = CreatePowerPoint();
        source.OpenXmlDocument.PresentationPart!.Presentation!.DefaultTextStyle =
            new DefaultTextStyle(new A.Level1ParagraphProperties { LeftMargin = 914400 });

        AssertPowerPointLoss(source, "presentation-settings");
    }

    [Fact]
    public void PictureBulletIsReportedBeforeStrictConversion() {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].AddTextBoxPoints("Item", 20, 20, 200, 40);
        A.Paragraph paragraph = source.OpenXmlDocument.PresentationPart!.SlideParts.Single()
            .Slide!.Descendants<A.Paragraph>().Single();
        paragraph.ParagraphProperties = new A.ParagraphProperties(
            new A.PictureBullet(new A.Blip { Embed = "rId999" }));

        AssertPowerPointLoss(source, "paragraph-layout");
    }

    [Fact]
    public void PlainNotesDrawingIsReportedBeforeStrictConversion() {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].Notes.Text = "Speaker note";
        source.Slides[0].AddRectangleCm(1, 1, 2, 1);
        var slidePart = source.OpenXmlDocument.PresentationPart!.SlideParts.Single();
        Shape drawing = slidePart.Slide!.Descendants<Shape>().Single();
        slidePart.NotesSlidePart!.NotesSlide!.CommonSlideData!.ShapeTree!
            .Append((Shape)drawing.CloneNode(true));

        AssertPowerPointLoss(source, "notes-slide-appearance");
    }

    [Fact]
    public void NotesConnectorIsReportedBeforeStrictConversion() {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].Notes.Text = "Speaker note";
        var slidePart = source.OpenXmlDocument.PresentationPart!.SlideParts.Single();
        slidePart.NotesSlidePart!.NotesSlide!.CommonSlideData!.ShapeTree!
            .Append(new ConnectionShape(
                new NonVisualConnectionShapeProperties(
                    new NonVisualDrawingProperties { Id = 200, Name = "Note connector" },
                    new NonVisualConnectorShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(new A.PresetGeometry(new A.AdjustValueList()) {
                    Preset = A.ShapeTypeValues.Line
                })));

        AssertPowerPointLoss(source, "notes-slide-appearance");
    }

    [Fact]
    public void UnusedSolidOdpMasterBackgroundIsReported() {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide();
        source.MasterPages[0].BackgroundColor = OdfColor.Parse("#336699");
        slide.MasterPageName = null;

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.ForFeature("masters-layouts"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void MalformedUnusedOdpMasterColorRemainsReportedWithoutThrowing() {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide();
        source.MasterPages[0].BackgroundColor = OdfColor.Parse("#336699");
        slide.MasterPageName = null;
        var styles = source.Package.GetXml("styles.xml");
        styles.Descendants(OdfNamespaces.Style + "drawing-page-properties").Single()
            .SetAttributeValue(OdfNamespaces.Draw + "fill-color", "invalid");
        source.Package.MarkXmlDirty("styles.xml");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.ForFeature("masters-layouts"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Approximated);
    }

    [Fact]
    public void SolidOdpFillWithTransparencyGradientIsReported() {
        OdpPresentation source = OdpPresentation.Create();
        OdpRectangle rectangle = source.AddSlide().AddRectangle(OdfRect.FromCentimeters(1, 1, 3, 2));
        rectangle.FillColor = OdfColor.Parse("#336699");
        string styleName = (string)rectangle.Element.Attribute(OdfNamespaces.Draw + "style-name")!;
        var style = source.Package.GetXml("content.xml").Descendants(OdfNamespaces.Style + "style")
            .Single(element => (string?)element.Attribute(OdfNamespaces.Style + "name") == styleName);
        style.Element(OdfNamespaces.Style + "graphic-properties")!
            .SetAttributeValue(OdfNamespaces.Draw + "fill-transparency-gradient-name", "Fade");
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.ForFeature("shape-appearance"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    private static PowerPointPresentation CreatePowerPoint() {
        PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        return source;
    }

    private static void AssertPowerPointLoss(PowerPointPresentation source, string feature) {
        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.ForFeature(feature), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }
}
