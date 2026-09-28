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

public sealed class PowerPointOdpStrictLossRegressionTests {
    private static readonly XNamespace Draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    private static readonly XNamespace Presentation = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
    private static readonly XNamespace Style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    private static readonly XNamespace Fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
    private static readonly XNamespace Table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";

    [Fact]
    public void HiddenOdpMasterBackgroundDoesNotBecomeSlideFill() {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide();
        source.MasterPages[0].BackgroundColor = OdfColor.Parse("#336699");
        source.Package.GetXml("content.xml").Descendants(Draw + "page").Single()
            .SetAttributeValue(Presentation + "background-visible", "false");
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> result = source.ToPowerPointPresentationResult();
        using PowerPointPresentation converted = result.Value;
        Assert.Null(converted.Slides[0].GetBackground().Color);
        AssertLoss(source, result, "slide-backgrounds");
    }

    [Fact]
    public void OdpGraphicPaddingIsReportedAsUnmappedAppearance() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddRectangle(OdfRect.FromCentimeters(1, 1, 4, 3)).FillColor = OdfColor.Parse("#336699");
        source.Package.GetXml("content.xml").Descendants(Style + "graphic-properties").Single()
            .SetAttributeValue(Fo + "padding-left", "0.2cm");
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> result = source.ToPowerPointPresentationResult();
        using PowerPointPresentation converted = result.Value;
        AssertLoss(source, result, "shape-appearance");
    }

    [Fact]
    public void SemiTransparentPowerPointSlideBackgroundIsNotMappedAsOpaque() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        var color = new A.RgbColorModelHex { Val = "336699" };
        color.Append(new A.Alpha { Val = 50000 });
        source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.Background =
            new Background(new BackgroundProperties(new A.SolidFill(color)));
        Assert.Equal(8, source.Slides[0].GetBackground().Color!.Length);

        OdfConversionResult<OdpPresentation> result = source.ToOpenDocumentResult();
        Assert.Null(result.Value.Slides[0].BackgroundColor);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "slide-backgrounds" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void OdpNoteTableIsReportedEvenWhenNotesContainNoPlainParagraph() {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide();
        slide.GetOrCreateSpeakerNotes();
        source.Package.GetXml("content.xml").Descendants(Presentation + "notes").Single()
            .Add(new XElement(Table + "table", new XAttribute(Table + "name", "NoteTable")));
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> result = source.ToPowerPointPresentationResult();
        using PowerPointPresentation converted = result.Value;
        AssertLoss(source, result, "speaker-notes");
    }

    [Fact]
    public void DescendingOdpLineReportsChangedGeometry() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddLine(OdfLength.Centimeters(5), OdfLength.Centimeters(1),
            OdfLength.Centimeters(1), OdfLength.Centimeters(5));

        OdfConversionResult<PowerPointPresentation> result = source.ToPowerPointPresentationResult();
        using PowerPointPresentation converted = result.Value;
        AssertLoss(source, result, "shape-transforms");
    }

    [Fact]
    public void ExplicitOdpZIndexIsReportedWhenXmlOrderWouldChange() {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide();
        slide.AddRectangle(OdfRect.FromCentimeters(1, 1, 4, 3));
        slide.AddRectangle(OdfRect.FromCentimeters(2, 2, 4, 3));
        XElement[] shapes = source.Package.GetXml("content.xml").Descendants(Draw + "rect").ToArray();
        shapes[0].SetAttributeValue(Draw + "z-index", "2");
        shapes[1].SetAttributeValue(Draw + "z-index", "1");
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> result = source.ToPowerPointPresentationResult();
        using PowerPointPresentation converted = result.Value;
        AssertLoss(source, result, "shapes");
    }

    [Fact]
    public void OdpAutomaticAdvanceWithoutTransitionStyleIsReported() {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide();
        slide.BackgroundColor = OdfColor.Parse("#336699");
        source.Package.GetXml("content.xml").Descendants(Style + "drawing-page-properties").Single()
            .SetAttributeValue(Presentation + "transition-change", "automatic");
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> result = source.ToPowerPointPresentationResult();
        using PowerPointPresentation converted = result.Value;
        AssertLoss(source, result, "slide-transitions");
    }

    private static void AssertLoss(OdpPresentation source,
        OdfConversionResult<PowerPointPresentation> result, string feature) {
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == feature &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }
}
