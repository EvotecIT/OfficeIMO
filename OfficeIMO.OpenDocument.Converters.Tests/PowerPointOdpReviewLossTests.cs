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
        using PowerPointDocument source = PowerPointDocument.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        PowerPointSlide slide = source.AddSlide(PowerPointSlideLayoutType.Blank);
        source.AddCustomShow("Executive path", new[] { slide });
        PowerPointAutoShape shape = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(1200000), OfficeIMO.OfficeLength.Emus(500000));
        slide.AddClassicAnimation(shape, PowerPointClassicAnimationEffect.Wipe);

        OdfConversionResult<OdpPresentation> result = source.ConvertToOpenDocumentWithEvidence();
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-custom-shows" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-classic-animations" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ConvertToOpenDocumentWithEvidence(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void AuthoredNotesMasterAppearanceRemainsStrictConversionLoss() {
        using PowerPointDocument source = PowerPointDocument.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        PowerPointSlide slide = source.AddSlide(PowerPointSlideLayoutType.Blank);
        slide.Notes.Text = "Speaker text";
        source.OpenXmlDocument.PresentationPart!.NotesMasterPart!.NotesMaster!
            .CommonSlideData!.Background = new Background(new BackgroundProperties(
                new A.SolidFill(new A.RgbColorModelHex { Val = "336699" })));

        OdfConversionResult<OdpPresentation> result = source.ConvertToOpenDocumentWithEvidence();
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "notes-master" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ConvertToOpenDocumentWithEvidence(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void ChangedNotesMasterPlaceholderGeometryRemainsStrictConversionLoss() {
        using PowerPointDocument source = PowerPointDocument.Create(new MemoryStream(),
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

        OdfConversionResult<OdpPresentation> result = source.ConvertToOpenDocumentWithEvidence();
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "notes-master" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void EditableReviewCommentsRemainStrictConversionLoss() {
        using PowerPointDocument source = PowerPointDocument.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        PowerPointSlide slide = source.AddSlide(PowerPointSlideLayoutType.Blank);
        source.AddClassicComment(slide, new PowerPointCommentAuthor("Reviewer"), "Please revise");

        OdfConversionResult<OdpPresentation> result = source.ConvertToOpenDocumentWithEvidence();
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-comments" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ConvertToOpenDocumentWithEvidence(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void PowerPointMetadataMapsCommonFieldsAndReportsOtherProperties() {
        using PowerPointDocument source = PowerPointDocument.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        source.Properties.Title = "Quarterly slides";
        source.Properties.Creator = "Ada";
        source.Properties.Subject = "Forecast";
        source.Properties.Description = "Board review";
        source.Properties.Keywords = "private,forecast";

        OdfConversionResult<OdpPresentation> result = source.ConvertToOpenDocumentWithEvidence();
        Assert.Equal("Quarterly slides", result.Value.Properties.Title);
        Assert.Equal("Ada", result.Value.Properties.Creator);
        Assert.Equal("Forecast", result.Value.Properties.Subject);
        Assert.Equal("Board review", result.Value.Properties.Description);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "document-metadata" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ConvertToOpenDocumentWithEvidence(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void SlideNamesRoundTripInBothDirections() {
        using PowerPointDocument source = PowerPointDocument.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank).Name = "Quarterly forecast";
        OdpPresentation odp = source.ToOpenDocument();
        Assert.Equal("Quarterly forecast", odp.Slides[0].Name);
        using PowerPointDocument reopened = odp.ToPowerPoint();
        Assert.Equal("Quarterly forecast", reopened.Slides[0].Name);
    }

    [Fact]
    public void DuplicatePowerPointSlideNamesAreUniquifiedWithLoss() {
        using PowerPointDocument source = PowerPointDocument.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank).Name = "Forecast";
        source.AddSlide(PowerPointSlideLayoutType.Blank).Name = "Forecast";
        OdfConversionResult<OdpPresentation> result = source.ConvertToOpenDocumentWithEvidence();
        Assert.Equal("Forecast", result.Value.Slides[0].Name);
        Assert.Equal("Forecast_2", result.Value.Slides[1].Name);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "slide-names" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ConvertToOpenDocumentWithEvidence(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void PackageLanguageMapsAndOtherCoreFieldsReportLoss() {
        using PowerPointDocument source = PowerPointDocument.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        source.OpenXmlDocument.PackageProperties.Language = "en-US";
        source.OpenXmlDocument.PackageProperties.Identifier = "urn:example:presentation";
        OdfConversionResult<OdpPresentation> result = source.ConvertToOpenDocumentWithEvidence();
        Assert.Equal("en-US", result.Value.Properties.Language);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "document-metadata" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ConvertToOpenDocumentWithEvidence(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Theory]
    [InlineData("animateMotion")]
    [InlineData("animateColor")]
    [InlineData("animateTransform")]
    [InlineData("set")]
    [InlineData("transitionFilter")]
    [InlineData("command")]
    [InlineData("audio")]
    public void OdpAnimationActionsRemainStrictConversionLoss(string action) {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide("Animated");
        slide.Element.Add(new XElement(OdfNamespaces.Anim + "par",
            new XElement(OdfNamespaces.Anim + action)));
        source.Package.MarkXmlDirty("content.xml");
        OdfConversionResult<PowerPointDocument> result = source.ConvertToPowerPointWithEvidence();
        using PowerPointDocument target = result.Value;
        Assert.Contains(result.Report.Mappings, mapping =>
            mapping.Feature == "source-presentation-animations" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ConvertToPowerPointWithEvidence(
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
        OdfConversionResult<PowerPointDocument> result = source.ConvertToPowerPointWithEvidence();
        using PowerPointDocument target = result.Value;
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ConvertToPowerPointWithEvidence(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void AuthoredOdpMasterAttributesReportLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide("Named");
        source.Package.GetXml("styles.xml")
            .Descendants(OdfNamespaces.Style + "master-page").Single()
            .SetAttributeValue(OdfNamespaces.Style + "display-name", "Corporate template");
        source.Package.MarkXmlDirty("styles.xml");
        OdfConversionResult<PowerPointDocument> result = source.ConvertToPowerPointWithEvidence();
        using PowerPointDocument target = result.Value;
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
    }

    [Fact]
    public void AuthoredOdpPageLayoutMarginRemainsStrictLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide("One");
        source.Package.GetXml("styles.xml")
            .Descendants(OdfNamespaces.Style + "page-layout-properties").Single()
            .SetAttributeValue(OdfNamespaces.Fo + "margin", "1cm");
        source.Package.MarkXmlDirty("styles.xml");
        OdfConversionResult<PowerPointDocument> result = source.ConvertToPowerPointWithEvidence();
        using PowerPointDocument target = result.Value;
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ConvertToPowerPointWithEvidence(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void AuthoredPowerPointMasterNameRemainsStrictLoss() {
        using PowerPointDocument source = PowerPointDocument.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        source.OpenXmlDocument.PresentationPart!.SlideMasterParts.Single().SlideMaster!
            .CommonSlideData!.Name = "Corporate";
        OdfConversionResult<OdpPresentation> result = source.ConvertToOpenDocumentWithEvidence();
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ConvertToOpenDocumentWithEvidence(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void AuthoredMasterTextStylesReportLossWithoutSlideText() {
        using PowerPointDocument source = PowerPointDocument.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        source.OpenXmlDocument.PresentationPart!.SlideMasterParts.Single().SlideMaster!.TextStyles!
            .ChildElements.First().Append(new A.Level1ParagraphProperties(
                new A.DefaultRunProperties { Bold = true }));
        OdfConversionResult<OdpPresentation> result = source.ConvertToOpenDocumentWithEvidence();
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
    }

    [Fact]
    public void AuthoredPowerPointMasterColorMapReportsLoss() {
        using PowerPointDocument source = PowerPointDocument.Create(new MemoryStream(),
            new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        source.OpenXmlDocument.PresentationPart!.SlideMasterParts.Single()
            .SlideMaster!.ColorMap!.Accent1 = A.ColorSchemeIndexValues.Accent2;
        OdfConversionResult<OdpPresentation> result = source.ConvertToOpenDocumentWithEvidence();
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ConvertToOpenDocumentWithEvidence(
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

        OdfConversionResult<PowerPointDocument> result = source.ConvertToPowerPointWithEvidence();
        using PowerPointDocument target = result.Value;
        if (hideMaster) {
            Assert.Null(target.Slides[0].BackgroundColor);
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "slide-backgrounds" &&
                mapping.Status == OdfConversionMappingStatus.Unsupported);
            Assert.Throws<OdfConversionLossException>(() => source.ConvertToPowerPointWithEvidence(
                new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
        } else {
            Assert.Equal("336699", target.Slides[0].BackgroundColor);
        }
    }
}
