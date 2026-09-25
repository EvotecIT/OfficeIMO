using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using A = DocumentFormat.OpenXml.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class PowerPointOdpPresentationLossCoverageTests {
    [Fact]
    public void PicturePlaceholderIsExplicitLoss() {
        using PowerPointPresentation source = CreateBlankPowerPoint();
        byte[] png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        using var image = new MemoryStream(png, writable: false);
        source.Slides[0].AddPicture(image, OfficeImageFormat.Png);
        Picture picture = source.OpenXmlDocument.PresentationPart!.SlideParts.First().Slide!
            .Descendants<Picture>().Single();
        picture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!
            .Append(new PlaceholderShape { Type = PlaceholderValues.Picture });

        AssertLoss(source, "placeholder-roles");
    }

    [Fact]
    public void NonTextShapePlaceholderIsExplicitLoss() {
        using PowerPointPresentation source = CreateBlankPowerPoint();
        source.Slides[0].AddRectanglePoints(20, 20, 100, 50);
        Shape shape = source.OpenXmlDocument.PresentationPart!.SlideParts.First().Slide!
            .Descendants<Shape>().Single();
        shape.TextBody?.Remove();
        shape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .Append(new PlaceholderShape { Type = PlaceholderValues.Object });

        AssertLoss(source, "placeholder-roles");
    }

    [Fact]
    public void TextPlaceholderIndexAndOrientationAreExplicitLoss() {
        using PowerPointPresentation source = CreateBlankPowerPoint();
        source.Slides[0].AddTextBoxPoints("Agenda", 20, 20, 100, 50);
        Shape shape = source.OpenXmlDocument.PresentationPart!.SlideParts.First().Slide!
            .Descendants<Shape>().Single();
        PlaceholderShape placeholder = shape.NonVisualShapeProperties!
            .ApplicationNonVisualDrawingProperties!.GetFirstChild<PlaceholderShape>()
            ?? shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties
                .AppendChild(new PlaceholderShape { Type = PlaceholderValues.Body });
        placeholder.Index = 7U;
        placeholder.Orientation = DirectionValues.Vertical;

        Assert.NotNull(shape.TextBody);
        Assert.Equal(7U, shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties
            .GetFirstChild<PlaceholderShape>()?.Index?.Value);

        AssertLoss(source, "placeholder-metadata");
    }

    [Fact]
    public void ThemeRunColorIsExplicitLoss() {
        using PowerPointPresentation source = CreateBlankPowerPoint();
        source.Slides[0].AddTextBoxPoints("Theme", 20, 20, 100, 50);
        A.Run run = source.OpenXmlDocument.PresentationPart!.SlideParts.First().Slide!
            .Descendants<A.Run>().Single();
        run.RunProperties ??= new A.RunProperties();
        run.RunProperties.Append(new A.SolidFill(new A.SchemeColor {
            Val = A.SchemeColorValues.Accent1 }));

        AssertLoss(source, "text-colors");
    }

    [Fact]
    public void TransformedDefaultRunColorIsExplicitLoss() {
        using PowerPointPresentation source = CreateBlankPowerPoint();
        source.Slides[0].AddTextBoxPoints("Alpha", 20, 20, 100, 50);
        A.Paragraph paragraph = source.OpenXmlDocument.PresentationPart!.SlideParts.First().Slide!
            .Descendants<A.Paragraph>().Single();
        paragraph.ParagraphProperties ??= new A.ParagraphProperties();
        paragraph.ParagraphProperties.Append(new A.DefaultRunProperties(
            new A.SolidFill(new A.RgbColorModelHex(new A.Alpha { Val = 50000 }) {
                Val = "336699" })));

        AssertLoss(source, "text-colors");
    }

    [Fact]
    public void PlainRgbParagraphDefaultColorIsExplicitLoss() {
        using PowerPointPresentation source = CreateBlankPowerPoint();
        source.Slides[0].AddTextBoxPoints("Inherited", 20, 20, 100, 50);
        A.Paragraph paragraph = source.OpenXmlDocument.PresentationPart!.SlideParts.First().Slide!
            .Descendants<A.Paragraph>().Single();
        paragraph.ParagraphProperties ??= new A.ParagraphProperties();
        paragraph.ParagraphProperties.Append(new A.DefaultRunProperties(
            new A.SolidFill(new A.RgbColorModelHex { Val = "336699" })));

        AssertLoss(source, "text-colors");
    }

    [Fact]
    public void ThemeHighlightAndSpeakerNoteColorAreExplicitLoss() {
        using PowerPointPresentation source = CreateBlankPowerPoint();
        source.Slides[0].AddTextBoxPoints("Highlight", 20, 20, 100, 50);
        A.Run run = source.OpenXmlDocument.PresentationPart!.SlideParts.First().Slide!
            .Descendants<A.Run>().Single();
        run.RunProperties ??= new A.RunProperties();
        run.RunProperties.Append(new A.Highlight(new A.SchemeColor {
            Val = A.SchemeColorValues.Accent1 }));
        AssertLoss(source, "text-colors");

        run.RunProperties.RemoveAllChildren<A.Highlight>();
        source.Slides[0].Notes.Text = "Speaker note";
        A.Run noteRun = source.OpenXmlDocument.PresentationPart!.SlideParts.First()
            .NotesSlidePart!.NotesSlide!.Descendants<A.Run>().Single();
        noteRun.RunProperties ??= new A.RunProperties();
        noteRun.RunProperties.Append(new A.SolidFill(new A.SchemeColor {
            Val = A.SchemeColorValues.Accent1 }));
        AssertLoss(source, "text-colors");
    }

    [Fact]
    public void AuthoredHandoutMasterIsExplicitLoss() {
        using PowerPointPresentation source = CreateBlankPowerPoint();
        PresentationPart presentation = source.OpenXmlDocument.PresentationPart!;
        HandoutMasterPart handout = presentation.HandoutMasterPart ?? presentation.AddNewPart<HandoutMasterPart>();
        CommonSlideData common = (CommonSlideData)presentation.SlideMasterParts.First()
            .SlideMaster!.CommonSlideData!.CloneNode(true);
        common.Background = new Background(new BackgroundProperties(
            new A.SolidFill(new A.RgbColorModelHex { Val = "336699" })));
        handout.HandoutMaster = new HandoutMaster(common);

        AssertLoss(source, "handout-master");
    }

    [Fact]
    public void AuthoredShowPropertiesAreExplicitLoss() {
        using PowerPointPresentation source = CreateBlankPowerPoint();
        PresentationPart presentation = source.OpenXmlDocument.PresentationPart!;
        PresentationPropertiesPart properties = presentation.PresentationPropertiesPart ??
            presentation.AddNewPart<PresentationPropertiesPart>();
        properties.PresentationProperties = new PresentationProperties(
            new ShowProperties { Loop = true });

        AssertLoss(source, "slide-show-settings");
    }

    [Fact]
    public void AuthoredViewPropertiesAreExplicitLossWithoutFlaggingStockDefaults() {
        using PowerPointPresentation source = CreateBlankPowerPoint();
        OdfConversionResult<OdpPresentation> baseline = source.ToOpenDocumentResult();
        Assert.DoesNotContain(baseline.Report.Mappings, mapping => mapping.Feature == "view-settings");
        source.OpenXmlDocument.PresentationPart!.ViewPropertiesPart!.ViewProperties!
            .SlideViewProperties!.CommonSlideViewProperties!.SnapToGrid = true;

        AssertLoss(source, "view-settings");
    }

    [Fact]
    public void MasterHeaderFooterBehaviorIsExplicitLoss() {
        using PowerPointPresentation source = CreateBlankPowerPoint();
        source.OpenXmlDocument.PresentationPart!.SlideMasterParts.First().SlideMaster!
            .Append(new HeaderFooter { DateTime = true });

        AssertLoss(source, "masters-layouts");
    }

    [Fact]
    public void ShapeDescriptionIsExplicitAccessibilityLoss() {
        using PowerPointPresentation source = CreateBlankPowerPoint();
        source.Slides[0].AddRectanglePoints(20, 20, 100, 50);
        Shape shape = source.OpenXmlDocument.PresentationPart!.SlideParts.First().Slide!
            .Descendants<Shape>().Single();
        shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Description = "Forecast chart";

        AssertLoss(source, "shape-accessibility");
    }

    [Fact]
    public void DecorativeShapeIsExplicitAccessibilityLoss() {
        using PowerPointPresentation source = CreateBlankPowerPoint();
        source.Slides[0].AddRectanglePoints(20, 20, 100, 50).Decorative = true;

        AssertLoss(source, "shape-accessibility");
    }

    [Fact]
    public void UnrelatedNonVisualExtensionDoesNotReportAccessibilityLoss() {
        using PowerPointPresentation source = CreateBlankPowerPoint();
        source.Slides[0].AddRectanglePoints(20, 20, 100, 50);
        Shape shape = source.OpenXmlDocument.PresentationPart!.SlideParts.First().Slide!
            .Descendants<Shape>().Single();
        shape.NonVisualShapeProperties!.NonVisualDrawingProperties!
            .Append(new A.NonVisualDrawingPropertiesExtensionList());

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.DoesNotContain(conversion.Report.Mappings, mapping =>
            mapping.Feature == "shape-accessibility");
    }

    [Fact]
    public void OdpCustomShowIsExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide("First");
        XDocument content = source.Package.GetXml("content.xml");
        XElement presentation = content.Descendants(OdfNamespaces.Office + "presentation").Single();
        presentation.Add(new XElement(OdfNamespaces.Presentation + "settings",
            new XElement(OdfNamespaces.Presentation + "show",
                new XAttribute(OdfNamespaces.Presentation + "name", "Tour"))));
        source.Package.MarkXmlDirty("content.xml");

        AssertLoss(source, "custom-shows");
    }

    [Fact]
    public void OdpPlaybackSettingsWithoutNamedShowAreExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide("First");
        XElement presentation = source.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Office + "presentation").Single();
        presentation.Add(new XElement(OdfNamespaces.Presentation + "settings",
            new XAttribute(OdfNamespaces.Presentation + "endless", "true")));
        source.Package.MarkXmlDirty("content.xml");

        AssertLoss(source, "slide-show-settings");
    }

    [Fact]
    public void OdpTypedTableValueIsExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddTable(OdfRect.FromCentimeters(1, 1, 8, 3), 1, 1);
        XElement cell = source.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table-cell").Single();
        cell.SetAttributeValue(OdfNamespaces.Office + "value-type", "float");
        cell.SetAttributeValue(OdfNamespaces.Office + "value", "12.5");
        source.Package.MarkXmlDirty("content.xml");

        AssertLoss(source, "table-values");
    }

    private static PowerPointPresentation CreateBlankPowerPoint() {
        PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        return source;
    }

    private static void AssertLoss(PowerPointPresentation source, string feature) {
        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == feature &&
            mapping.Status != OdfConversionMappingStatus.Converted);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    private static void AssertLoss(OdpPresentation source, string feature) {
        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == feature &&
            mapping.Status != OdfConversionMappingStatus.Converted);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }
}
