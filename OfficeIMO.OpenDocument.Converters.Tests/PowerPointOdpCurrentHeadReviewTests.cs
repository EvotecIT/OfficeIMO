using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using A = DocumentFormat.OpenXml.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class PowerPointOdpCurrentHeadReviewTests {
    [Fact]
    public void CustomOdpGeneratorIsMetadataLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide();
        source.Metadata.Generator = "External editor";

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "document-metadata" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Theory]
    [InlineData("Application")]
    [InlineData("ApplicationVersion")]
    [InlineData("PresentationFormat")]
    [InlineData("TotalTime")]
    public void AuthoredPowerPointExtendedMetadataIsLoss(string property) {
        using PowerPointPresentation source = CreatePowerPoint();
        switch (property) {
            case "Application": source.ApplicationProperties.Application = "External editor"; break;
            case "ApplicationVersion": source.ApplicationProperties.ApplicationVersion = "2.0"; break;
            case "PresentationFormat": source.ApplicationProperties.PresentationFormat = "Custom"; break;
            case "TotalTime": source.ApplicationProperties.TotalTime = "45"; break;
        }

        AssertPowerPointLoss(source, "document-metadata");
    }

    [Fact]
    public void AuthoredPowerPointHyperlinkBaseIsMetadataLoss() {
        using PowerPointPresentation source = CreatePowerPoint();
        source.OpenXmlDocument.ExtendedFilePropertiesPart!.Properties!
            .AppendChild(new DocumentFormat.OpenXml.ExtendedProperties.HyperlinkBase("https://example.test/slides/"));

        AssertPowerPointLoss(source, "document-metadata");
    }

    [Theory]
    [InlineData("FirstSlide")]
    [InlineData("RightToLeft")]
    [InlineData("TitlePlaceholder")]
    [InlineData("NotesSize")]
    public void AuthoredPowerPointRootSettingsAreLoss(string setting) {
        using PowerPointPresentation source = CreatePowerPoint();
        Presentation root = source.OpenXmlDocument.PresentationPart!.Presentation!;
        switch (setting) {
            case "FirstSlide": root.FirstSlideNum = 4; break;
            case "RightToLeft": root.RightToLeft = true; break;
            case "TitlePlaceholder": root.ShowSpecialPlaceholderOnTitleSlide = false; break;
            case "NotesSize": root.NotesSize = new NotesSize { Cx = 5_000_000, Cy = 7_000_000 }; break;
        }

        AssertPowerPointLoss(source, "presentation-settings");
    }

    [Fact]
    public void StockPowerPointRootAndApplicationMetadataAreNotLoss() {
        using PowerPointPresentation source = CreatePowerPoint();
        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.DoesNotContain(conversion.Report.Mappings, mapping =>
            (mapping.Feature == "presentation-settings" || mapping.Feature == "document-metadata") &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void ExplicitDefaultPowerPointRootSettingsAreNotLoss() {
        using PowerPointPresentation source = CreatePowerPoint();
        Presentation root = source.OpenXmlDocument.PresentationPart!.Presentation!;
        root.FirstSlideNum = 1;
        root.RightToLeft = false;
        root.ShowSpecialPlaceholderOnTitleSlide = true;

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "presentation-settings" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void UnfilledPowerPointLineIsNotShapeAppearanceLoss() {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].AddLine(0, 0, 100000, 100000);
        ShapeProperties properties = source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<Shape>().Single().ShapeProperties!;
        Assert.NotNull(properties.GetFirstChild<A.NoFill>());

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "shape-appearance" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Theory]
    [InlineData("spd", "fast")]
    [InlineData("dur", "500")]
    public void TransitionSpeedAndDurationRemainExplicitLoss(string attribute, string value) {
        using PowerPointPresentation source = CreatePowerPoint();
        Slide slide = source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!;
        slide.Transition = new Transition();
        slide.Transition.SetAttribute(new OpenXmlAttribute("", attribute, "", value));

        AssertPowerPointLoss(source, "slide-transition-timing");
    }

    [Fact]
    public void TextBodyLevelParagraphDefaultsRemainExplicitLoss() {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].AddTextBoxPoints("Level text", 20, 20, 200, 40);
        Shape shape = source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<Shape>().Single();
        A.ListStyle listStyle = shape.TextBody!.GetFirstChild<A.ListStyle>()
            ?? shape.TextBody.AppendChild(new A.ListStyle());
        listStyle.Append(new A.Level1ParagraphProperties { LeftMargin = 457200 });

        AssertPowerPointLoss(source, "paragraph-layout");
    }

    [Fact]
    public void DirectOdpImageCropMapsWithoutShapeAppearanceLoss() {
        OdpPresentation source = OdpPresentation.Create();
        byte[] png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        OdpImage image = source.AddSlide().AddImage(png, "pixel.png", OdfRect.FromCentimeters(1, 1, 2, 2));
        image.Crop = new OdfInsets(OdfLength.Centimeters(0), OdfLength.Centimeters(0),
            OdfLength.Centimeters(0), OdfLength.Centimeters(0.2));

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.InRange(Assert.Single(target.Slides[0].Pictures).CropLeftRatio, 0.09, 0.11);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "shape-appearance" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void NestedOdpTableListRemainsExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddTable(OdfRect.FromCentimeters(1, 1, 8, 3), 1, 1);
        XElement cell = source.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table-cell").Single();
        cell.Add(new XElement(OdfNamespaces.Text + "list",
            new XElement(OdfNamespaces.Text + "list-item",
                new XElement(OdfNamespaces.Text + "p", "Nested item"))));
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "text-lists" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    private static PowerPointPresentation CreatePowerPoint() {
        PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        return source;
    }

    private static void AssertPowerPointLoss(PowerPointPresentation source, string feature) {
        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == feature &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }
}
