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

public sealed class PowerPointOdpReviewLossTests {
    [Fact]
    public void EditableCustomShowAndAnimationRemainStrictConversionLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        PowerPointSlide slide = source.AddSlide(PowerPointSlideLayoutType.Blank);
        source.AddCustomShow("Executive path", new[] { slide });
        PowerPointAutoShape shape = slide.AddRectangle(100000, 100000, 1200000, 500000);
        slide.AddClassicAnimation(shape, PowerPointClassicAnimationEffect.Wipe);

        OdfConversionResult<OdpPresentation> result = source.ToOpenDocumentResult();
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-custom-shows" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-classic-animations" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void AuthoredNotesMasterAppearanceRemainsStrictConversionLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        PowerPointSlide slide = source.AddSlide(PowerPointSlideLayoutType.Blank);
        slide.Notes.Text = "Speaker text";
        source.OpenXmlDocument.PresentationPart!.NotesMasterPart!.NotesMaster!
            .CommonSlideData!.Background = new Background(new BackgroundProperties(
                new A.SolidFill(new A.RgbColorModelHex { Val = "336699" })));

        OdfConversionResult<OdpPresentation> result = source.ToOpenDocumentResult();
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "notes-master" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void ChangedNotesMasterPlaceholderGeometryRemainsStrictConversionLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        PowerPointSlide slide = source.AddSlide(PowerPointSlideLayoutType.Blank);
        slide.Notes.Text = "Speaker text";
        source.OpenXmlDocument.PresentationPart!.NotesMasterPart!.NotesMaster!
            .CommonSlideData!.ShapeTree!.Append(new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = 2U, Name = "Notes body position" },
                    new NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = PlaceholderValues.Body })),
                new ShapeProperties(new A.Transform2D(
                    new A.Offset { X = 1000, Y = 0 },
                    new A.Extents { Cx = 914400, Cy = 914400 }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
                new TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph())));

        OdfConversionResult<OdpPresentation> result = source.ToOpenDocumentResult();
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "notes-master" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void EditableReviewCommentsRemainStrictConversionLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        PowerPointSlide slide = source.AddSlide(PowerPointSlideLayoutType.Blank);
        source.AddClassicComment(slide, new PowerPointCommentAuthor("Reviewer"), "Please revise");

        OdfConversionResult<OdpPresentation> result = source.ToOpenDocumentResult();
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-comments" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void PowerPointMetadataMapsCommonFieldsAndReportsOtherProperties() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        source.BuiltinDocumentProperties.Title = "Quarterly slides";
        source.BuiltinDocumentProperties.Creator = "Ada";
        source.BuiltinDocumentProperties.Subject = "Forecast";
        source.BuiltinDocumentProperties.Description = "Board review";
        source.BuiltinDocumentProperties.Keywords = "private,forecast";

        OdfConversionResult<OdpPresentation> result = source.ToOpenDocumentResult();
        Assert.Equal("Quarterly slides", result.Value.Metadata.Title);
        Assert.Equal("Ada", result.Value.Metadata.Creator);
        Assert.Equal("Forecast", result.Value.Metadata.Subject);
        Assert.Equal("Board review", result.Value.Metadata.Description);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "document-metadata" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void SlideNamesRoundTripInBothDirections() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank).Name = "Quarterly forecast";
        OdpPresentation odp = source.ToOpenDocument();
        Assert.Equal("Quarterly forecast", odp.Slides[0].Name);
        using PowerPointPresentation reopened = odp.ToPowerPointPresentation();
        Assert.Equal("Quarterly forecast", reopened.Slides[0].Name);
    }

    [Fact]
    public void PackageLanguageMapsAndOtherCoreFieldsReportLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        source.OpenXmlDocument.PackageProperties.Language = "en-US";
        source.OpenXmlDocument.PackageProperties.Identifier = "urn:example:presentation";
        OdfConversionResult<OdpPresentation> result = source.ToOpenDocumentResult();
        Assert.Equal("en-US", result.Value.Metadata.Language);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "document-metadata" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Theory]
    [InlineData("animateMotion")]
    [InlineData("animateColor")]
    [InlineData("animateTransform")]
    [InlineData("set")]
    [InlineData("transitionFilter")]
    public void OdpAnimationActionsRemainStrictConversionLoss(string action) {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide("Animated");
        slide.Element.Add(new XElement(OdfNamespaces.Anim + "par",
            new XElement(OdfNamespaces.Anim + action)));
        source.Package.MarkXmlDirty("content.xml");
        OdfConversionResult<PowerPointPresentation> result = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = result.Value;
        Assert.Contains(result.Report.Mappings, mapping =>
            mapping.Feature == "source-presentation-animations" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void AuthoredSingleOdpMasterAndLayoutNamesReportLoss() {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide("Named");
        XDocument styles = source.Package.GetXml("styles.xml");
        XElement master = styles.Descendants(OdfNamespaces.Style + "master-page").Single();
        XElement layout = styles.Descendants(OdfNamespaces.Style + "presentation-page-layout").Single();
        master.SetAttributeValue(OdfNamespaces.Style + "name", "Corporate");
        layout.SetAttributeValue(OdfNamespaces.Style + "name", "BoardLayout");
        slide.MasterPageName = "Corporate";
        slide.LayoutName = "BoardLayout";
        source.Package.MarkXmlDirty("styles.xml");
        OdfConversionResult<PowerPointPresentation> result = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = result.Value;
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void AuthoredPowerPointMasterColorMapReportsLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        source.OpenXmlDocument.PresentationPart!.SlideMasterParts.Single()
            .SlideMaster!.ColorMap!.Accent1 = A.ColorSchemeIndexValues.Accent2;
        OdfConversionResult<OdpPresentation> result = source.ToOpenDocumentResult();
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OdpDefaultDrawingPageBackgroundAndHiddenMasterAreResolved(bool hideMaster) {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide();
        source.MasterPages[0].BackgroundColor = OdfColor.Parse("#AA2200");
        slide.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", null);
        XElement properties = new(OdfNamespaces.Style + "drawing-page-properties");
        if (hideMaster)
            properties.SetAttributeValue(OdfNamespaces.Presentation + "background-visible", "false");
        else {
            properties.SetAttributeValue(OdfNamespaces.Draw + "fill", "solid");
            properties.SetAttributeValue(OdfNamespaces.Draw + "fill-color", "#336699");
        }
        source.Package.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(
            new XElement(OdfNamespaces.Style + "default-style",
                new XAttribute(OdfNamespaces.Style + "family", "drawing-page"), properties));
        source.Package.MarkXmlDirty("styles.xml");
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> result = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = result.Value;
        if (hideMaster) {
            Assert.Null(target.Slides[0].BackgroundColor);
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "slide-backgrounds" &&
                mapping.Status == OdfConversionMappingStatus.Unsupported);
            Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
                new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
        } else {
            Assert.Equal("336699", target.Slides[0].BackgroundColor);
        }
    }
}
