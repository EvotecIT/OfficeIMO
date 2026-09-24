using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using A = DocumentFormat.OpenXml.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class PowerPointOdpPresentationSemanticsTests {
    [Fact]
    public void PowerPointSectionsRemainExplicitLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        source.AddSection("Results", startSlideIndex: 0);

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "sections" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void AuthoredUnusedBlankLayoutMetadataAddsMasterLayoutLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Title);
        int before = source.ToOpenDocumentResult().Report.Mappings
            .Where(mapping => mapping.Feature == "masters-layouts")
            .Sum(mapping => mapping.Count);
        SlideLayout blank = source.OpenXmlDocument.PresentationPart!.SlideMasterParts.First()
            .SlideLayoutParts.Select(part => part.SlideLayout!)
            .Single(layout => layout.Type?.Value == SlideLayoutValues.Blank);
        blank.CommonSlideData!.Name = "Authored unused blank layout";

        int after = source.ToOpenDocumentResult().Report.Mappings
            .Where(mapping => mapping.Feature == "masters-layouts")
            .Sum(mapping => mapping.Count);
        Assert.True(after > before);
    }

    [Fact]
    public void AuthoredSelectedBlankLayoutMetadataIsExplicitLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        source.OpenXmlDocument.PresentationPart!.SlideParts.First().SlideLayoutPart!
            .SlideLayout!.CommonSlideData!.Name = "Authored selected blank layout";

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void OdpSpeakerNoteFrameGeometryIsExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().GetOrCreateSpeakerNotes().AddParagraph("Presenter note");
        XNamespace presentation = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        XNamespace svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
        XElement frame = source.Package.GetXml("content.xml").Descendants(presentation + "notes")
            .Descendants(draw + "frame").Single();
        frame.SetAttributeValue(svg + "x", "2cm");
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "speaker-notes" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void OdpSpeakerNotesPageStyleIsExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().GetOrCreateSpeakerNotes().AddParagraph("Presenter note");
        XNamespace presentation = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        source.Package.GetXml("content.xml").Descendants(presentation + "notes").Single()
            .SetAttributeValue(draw + "style-name", "NotesPageStyle");
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "speaker-notes" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void OdpSpeakerNotesInheritedDefaultFrameStyleIsExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().GetOrCreateSpeakerNotes().AddParagraph("Presenter note");
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        source.Package.GetXml("styles.xml").Root!.Element(office + "styles")!.Add(
            new XElement(style + "default-style", new XAttribute(style + "family", "graphic"),
                new XElement(style + "graphic-properties", new XAttribute(draw + "fill", "solid"),
                    new XAttribute(draw + "fill-color", "#336699"))));
        source.Package.MarkXmlDirty("styles.xml");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "speaker-notes" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void OdpNonTextPresentationClassIsExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddRectangle(OdfRect.FromCentimeters(1, 1, 4, 2));
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        XNamespace presentation = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
        source.Package.GetXml("content.xml").Descendants(draw + "rect").Single()
            .SetAttributeValue(presentation + "class", "graphic");
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "placeholder-roles" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void OdpDefaultGraphicStyleMapsDirectSolidFill() {
        OdpPresentation source = OdpPresentation.Create();
        OdpRectangle rectangle = source.AddSlide().AddRectangle(OdfRect.FromCentimeters(1, 1, 4, 2));
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        source.Package.GetXml("styles.xml").Root!.Element(office + "styles")!.Add(
            new XElement(style + "default-style", new XAttribute(style + "family", "graphic"),
                new XElement(style + "graphic-properties", new XAttribute(draw + "fill", "solid"),
                    new XAttribute(draw + "fill-color", "#336699"))));
        source.Package.MarkXmlDirty("styles.xml");

        Assert.Equal("#336699", rectangle.FillColor?.ToString());
        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.NotNull(target.Slides[0].Shapes.OfType<PowerPointAutoShape>().Single().FillColor);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "shape-appearance" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void TextBodyPictureEffectsAndTextGeometryHaveExplicitLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        PowerPointSlide slide = source.AddSlide(PowerPointSlideLayoutType.Blank);
        PowerPointTextBox box = slide.AddTextBoxPoints("Inset text", 20, 20, 200, 40);
        box.TextMarginLeftPoints = 18;
        slide.AddTextShape(OfficePresetShapeType.Ellipse, "Oval text");
        byte[] png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        using var image = new MemoryStream(png, writable: false);
        slide.AddPicture(image, OfficeImageFormat.Png).GrayScale = true;

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "shape-appearance" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "shape-geometry" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void UnusedAuthoredLayoutHasExplicitLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        var master = source.OpenXmlDocument.PresentationPart!.SlideMasterParts.First();
        var used = source.OpenXmlDocument.PresentationPart.SlideParts.Single().SlideLayoutPart;
        var unused = master.SlideLayoutParts.First(layout => !ReferenceEquals(layout, used));
        unused.SlideLayout!.CommonSlideData!.Background = new Background(new BackgroundProperties(
            new A.SolidFill(new A.RgbColorModelHex { Val = "336699" })));

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void UnusedLayoutPlaceholderGeometryHasExplicitLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        var master = source.OpenXmlDocument.PresentationPart!.SlideMasterParts.First();
        var unused = master.SlideLayoutParts.First(layout => layout.SlideLayout!
            .CommonSlideData!.ShapeTree!.Elements<Shape>().Any());
        Shape placeholder = unused.SlideLayout!.CommonSlideData!.ShapeTree!.Elements<Shape>().First();
        A.Offset offset = placeholder.ShapeProperties!.GetFirstChild<A.Transform2D>()!.Offset!;
        offset.X = offset.X!.Value + 1000L;

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
    }

    [Fact]
    public void InheritedGraphicStrokeWidthReachesPowerPoint() {
        OdpPresentation source = OdpPresentation.Create();
        OdpRectangle rectangle = source.AddSlide().AddRectangle(OdfRect.FromCentimeters(1, 1, 5, 3));
        rectangle.StrokeColor = OdfColor.Parse("#336699");
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XNamespace svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
        source.Package.GetXml("styles.xml").Root!.Element(office + "styles")!.Add(
            new XElement(style + "style", new XAttribute(style + "name", "WideStroke"),
                new XAttribute(style + "family", "graphic"),
                new XElement(style + "graphic-properties", new XAttribute(svg + "stroke-width", "3pt"))));
        XElement child = source.Package.GetXml("content.xml").Descendants(style + "style")
            .Single(item => (string?)item.Attribute(style + "family") == "graphic");
        child.SetAttributeValue(style + "parent-style-name", "WideStroke");
        source.Package.MarkXmlDirty("styles.xml");
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Equal(3, Assert.Single(target.Slides[0].Shapes.OfType<PowerPointAutoShape>()).OutlineWidthPoints);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "shape-appearance");
    }

    [Fact]
    public void InheritedSlideOpacityWithoutFillHasExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide();
        source.MasterPages[0].BackgroundColor = OdfColor.Parse("#336699");
        slide.BackgroundColor = OdfColor.Parse("#CC5500");
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        source.Package.GetXml("styles.xml").Root!.Element(office + "styles")!.Add(
            new XElement(style + "style", new XAttribute(style + "name", "FadedSlide"),
                new XAttribute(style + "family", "drawing-page"),
                new XElement(style + "drawing-page-properties", new XAttribute(draw + "opacity", "50%"))));
        XElement child = source.Package.GetXml("content.xml").Descendants(style + "style")
            .Single(item => (string?)item.Attribute(style + "family") == "drawing-page");
        child.SetAttributeValue(style + "parent-style-name", "FadedSlide");
        child.Element(style + "drawing-page-properties")!.Remove();
        source.Package.MarkXmlDirty("styles.xml");
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "slide-backgrounds" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void RawOdpCustomShapeHasExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide();
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        source.Package.GetXml("content.xml").Descendants(draw + "page").Single()
            .Add(new XElement(draw + "custom-shape", new XAttribute(draw + "name", "Star")));
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "shapes" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void GradientMasterWithStaleSolidColorDoesNotProjectColor() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide();
        source.MasterPages[0].BackgroundColor = OdfColor.Parse("#336699");
        XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        XElement properties = source.Package.GetXml("styles.xml")
            .Descendants(style + "drawing-page-properties").Single();
        properties.SetAttributeValue(draw + "fill", "gradient");
        source.Package.MarkXmlDirty("styles.xml");

        Assert.Null(source.MasterPages[0].BackgroundColor);
        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Null(target.Slides[0].GetBackground().Color);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
    }

    [Fact]
    public void InheritedOdpGradientStylesAreReportedForMasterAndShape() {
        OdpPresentation source = OdpPresentation.Create();
        OdpRectangle rectangle = source.AddSlide().AddRectangle(OdfRect.FromCentimeters(1, 1, 5, 3));
        source.MasterPages[0].BackgroundColor = OdfColor.Parse("#336699");
        rectangle.FillColor = OdfColor.Parse("#CC5500");
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        XDocument styles = source.Package.GetXml("styles.xml");
        XElement named = styles.Root!.Element(office + "styles")!;
        named.Add(new XElement(style + "style",
            new XAttribute(style + "name", "GradientMaster"), new XAttribute(style + "family", "drawing-page"),
            new XElement(style + "drawing-page-properties", new XAttribute(draw + "fill", "gradient"))));
        named.Add(new XElement(style + "style",
            new XAttribute(style + "name", "GradientShape"), new XAttribute(style + "family", "graphic"),
            new XElement(style + "graphic-properties", new XAttribute(draw + "fill", "gradient"))));
        XElement masterStyle = styles.Descendants(style + "style").Single(item =>
            (string?)item.Attribute(style + "family") == "drawing-page" &&
            (string?)item.Attribute(style + "name") != "GradientMaster");
        masterStyle.SetAttributeValue(style + "parent-style-name", "GradientMaster");
        masterStyle.Element(style + "drawing-page-properties")!.Remove();
        XDocument content = source.Package.GetXml("content.xml");
        XElement shapeStyle = content.Descendants(style + "style").Single(item =>
            (string?)item.Attribute(style + "family") == "graphic");
        shapeStyle.SetAttributeValue(style + "parent-style-name", "GradientShape");
        shapeStyle.Element(style + "graphic-properties")!.Remove();
        source.Package.MarkXmlDirty("styles.xml");
        source.Package.MarkXmlDirty("content.xml");

        Assert.Null(source.MasterPages[0].BackgroundColor);
        Assert.Null(rectangle.FillColor);
        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Null(target.Slides[0].GetBackground().Color);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "shape-appearance" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void SolidChildMasterDoesNotInheritInactiveGradientMetadata() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide();
        source.MasterPages[0].BackgroundColor = OdfColor.Parse("#336699");
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        XDocument styles = source.Package.GetXml("styles.xml");
        styles.Root!.Element(office + "styles")!.Add(new XElement(style + "style",
            new XAttribute(style + "name", "ParentGradient"),
            new XAttribute(style + "family", "drawing-page"),
            new XElement(style + "drawing-page-properties",
                new XAttribute(draw + "fill", "gradient"),
                new XAttribute(draw + "fill-gradient-name", "UnusedGradient"))));
        XElement child = styles.Descendants(style + "style").Single(item =>
            (string?)item.Attribute(style + "family") == "drawing-page" &&
            (string?)item.Attribute(style + "name") != "ParentGradient");
        child.SetAttributeValue(style + "parent-style-name", "ParentGradient");
        source.Package.MarkXmlDirty("styles.xml");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Equal("336699", target.Slides[0].GetBackground().Color);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "masters-layouts");
    }

    [Fact]
    public void EmptyMasterAndLayoutDoNotTriggerStrictLoss() {
        using PowerPointPresentation powerPoint = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        powerPoint.AddSlide(PowerPointSlideLayoutType.Blank);
        OdfConversionResult<OdpPresentation> fromPowerPoint = powerPoint.ToOpenDocumentResult();
        Assert.DoesNotContain(fromPowerPoint.Report.Mappings, mapping => mapping.Feature == "masters-layouts");
        OdpPresentation strictOdp = powerPoint.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }).Value;
        Assert.Single(strictOdp.Slides);

        OdpPresentation odf = OdpPresentation.Create();
        odf.AddSlide();
        OdfConversionResult<PowerPointPresentation> fromOdp = odf.ToPowerPointPresentationResult();
        using PowerPointPresentation converted = fromOdp.Value;
        Assert.DoesNotContain(fromOdp.Report.Mappings, mapping => mapping.Feature == "masters-layouts");
        using PowerPointPresentation strictPowerPoint = odf.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }).Value;
        Assert.Single(strictPowerPoint.Slides);
    }

    [Fact]
    public void ThemeMasterBackgroundReportsLostInheritance() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.OpenXmlDocument.PresentationPart!.SlideMasterParts.First().SlideMaster!
            .CommonSlideData!.Background = new Background(new BackgroundProperties(
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 })));
        source.AddSlide(PowerPointSlideLayoutType.Blank);

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void MasterTextStylesReportInheritedFormattingLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.OpenXmlDocument.PresentationPart!.SlideMasterParts.First().SlideMaster!.TextStyles!
            .ChildElements.First().Append(new A.Level1ParagraphProperties(
                new A.DefaultRunProperties { Bold = true }));
        source.AddSlide(PowerPointSlideLayoutType.Blank)
            .AddTextBoxPoints("Inherited text", 20, 20, 200, 40);

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
    }

    [Fact]
    public void EmptyMasterTextStyleContainersDoNotReportLossForOrdinaryText() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank)
            .AddTextBoxPoints("Ordinary text", 20, 20, 200, 40);

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "masters-layouts");
    }

    [Fact]
    public void OdpGradientMasterBackgroundReportsUnsupportedInheritance() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide();
        source.MasterPages[0].BackgroundColor = OdfColor.Parse("#336699");
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        XDocument styles = source.Package.GetXml("styles.xml");
        XElement properties = styles.Descendants(style + "drawing-page-properties").Single();
        properties.SetAttributeValue(draw + "fill", "gradient");
        properties.SetAttributeValue(draw + "fill-color", null);
        properties.SetAttributeValue(draw + "fill-gradient-name", "BackgroundGradient");
        styles.Root!.Element(office + "styles")!.Add(new XElement(draw + "gradient",
            new XAttribute(draw + "name", "BackgroundGradient"),
            new XAttribute(draw + "style", "linear"),
            new XAttribute(draw + "start-color", "#336699"),
            new XAttribute(draw + "end-color", "#99CCFF")));
        source.Package.MarkXmlDirty("styles.xml");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void MasterBackgroundAndDistinctLayoutsSurviveOdpReopen() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.OpenXmlDocument.PresentationPart!.SlideMasterParts.First().SlideMaster!
            .CommonSlideData!.Background = new Background(new BackgroundProperties(
                new A.SolidFill(new A.RgbColorModelHex { Val = "336699" })));
        source.AddSlide(PowerPointSlideLayoutType.Title);
        PowerPointSlide overrideSlide = source.AddSlide(PowerPointSlideLayoutType.Blank);
        overrideSlide.BackgroundColor = "CC5500";

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        OdpPresentation reopened = OdpPresentation.Load(new MemoryStream(conversion.Value.ToBytes()));
        Assert.True(reopened.Validate().IsValid);
        Assert.Single(reopened.MasterPages);
        Assert.Equal(2, reopened.Layouts.Count);
        Assert.Equal(reopened.Slides[0].MasterPageName, reopened.Slides[1].MasterPageName);
        Assert.NotEqual(reopened.Slides[0].LayoutName, reopened.Slides[1].LayoutName);
        Assert.Equal("#336699", reopened.MasterPages[0].BackgroundColor?.ToString());
        Assert.Null(reopened.Slides[0].BackgroundColor);
        Assert.Equal("#CC5500", reopened.Slides[1].BackgroundColor?.ToString());
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "master-associations" &&
            mapping.Status == OdfConversionMappingStatus.Converted && mapping.Count == 2);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "layout-associations" &&
            mapping.Status == OdfConversionMappingStatus.Converted && mapping.Count == 2);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);

        OdfConversionResult<PowerPointPresentation> back = reopened.ToPowerPointPresentationResult();
        using PowerPointPresentation converted = back.Value;
        using PowerPointPresentation reopenedPowerPoint = PowerPointPresentation.Load(new MemoryStream(converted.ToBytes()));
        Assert.Equal("336699", reopenedPowerPoint.Slides[0].GetBackground().Color);
        Assert.Equal("CC5500", reopenedPowerPoint.Slides[1].GetBackground().Color);
    }

    [Fact]
    public void TitleAndBodyPlaceholderRolesSurviveOdpAndPowerPointReopen() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        PowerPointSlide slide = source.AddSlide();
        PowerPointTextBox title = slide.AddTextBoxPoints("Quarterly results", 20, 20, 300, 40);
        title.PlaceholderType = PowerPointPlaceholderType.Title;
        PowerPointTextBox body = slide.AddTextBoxPoints("Revenue grew", 20, 75, 300, 80);
        body.PlaceholderType = PowerPointPlaceholderType.Body;

        OdfConversionResult<OdpPresentation> toOdp = source.ToOpenDocumentResult();
        Assert.Contains(toOdp.Report.Mappings, mapping => mapping.Feature == "placeholder-roles" &&
            mapping.Status == OdfConversionMappingStatus.Converted && mapping.Count == 2);
        OdpPresentation reopenedOdp = OdpPresentation.Load(new MemoryStream(toOdp.Value.ToBytes()));
        Assert.True(reopenedOdp.Validate().IsValid);
        Assert.Equal(new[] { "title", "outline" }, reopenedOdp.Slides[0].Shapes
            .OfType<OdpTextBox>().Select(box => box.PresentationClass));

        OdfConversionResult<PowerPointPresentation> toPowerPoint = reopenedOdp.ToPowerPointPresentationResult();
        using PowerPointPresentation converted = toPowerPoint.Value;
        Assert.Contains(toPowerPoint.Report.Mappings, mapping => mapping.Feature == "placeholder-roles" &&
            mapping.Status == OdfConversionMappingStatus.Converted && mapping.Count == 2);
        using PowerPointPresentation reopenedPowerPoint = PowerPointPresentation.Load(new MemoryStream(converted.ToBytes()));
        Assert.Equal(new PowerPointPlaceholderType?[] { PowerPointPlaceholderType.Title, PowerPointPlaceholderType.Body },
            reopenedPowerPoint.Slides[0].TextBoxes.Select(box => box.PlaceholderType));
    }

    [Fact]
    public void UnsupportedOdpPresentationClassIsExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2), "Chart title")
            .PresentationClass = "chart";

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "placeholder-roles" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void UnsupportedSlideOverrideDoesNotExposeMappedMasterColor() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.OpenXmlDocument.PresentationPart!.SlideMasterParts.First().SlideMaster!
            .CommonSlideData!.Background = new Background(new BackgroundProperties(
                new A.SolidFill(new A.RgbColorModelHex { Val = "336699" })));
        source.AddSlide(PowerPointSlideLayoutType.Blank).SetBackgroundGradient("FF0000", "0000FF");

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        OdpPresentation reopened = OdpPresentation.Load(new MemoryStream(conversion.Value.ToBytes()));
        Assert.True(reopened.Validate().IsValid);
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        Assert.Equal("#336699", reopened.MasterPages[0].BackgroundColor?.ToString());
        Assert.Contains(reopened.Package.GetXml("content.xml").Descendants(style + "drawing-page-properties"),
            properties => (string?)properties.Attribute(draw + "fill") == "none");
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "slide-backgrounds" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void OdpGradientSlideBackgroundIsExplicitLossWithoutMasterFallback() {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide();
        source.MasterPages[0].BackgroundColor = OdfColor.Parse("#336699");
        slide.BackgroundColor = OdfColor.Parse("#CC5500");
        XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        XElement properties = source.Package.GetXml("content.xml")
            .Descendants(style + "drawing-page-properties").Single();
        properties.SetAttributeValue(draw + "fill", "gradient");
        properties.SetAttributeValue(draw + "fill-color", null);
        properties.SetAttributeValue(draw + "fill-gradient-name", "Gradient1");
        source.Package.MarkXmlDirty("content.xml");

        Assert.Null(slide.BackgroundColor);
        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Null(target.Slides[0].GetBackground().Color);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "slide-backgrounds" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void AbsentPlaceholderTypeUsesObjectRole() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        PowerPointTextBox box = source.AddSlide(PowerPointSlideLayoutType.Blank)
            .AddTextBoxPoints("Object", 20, 20, 200, 40);
        box.PlaceholderType = PowerPointPlaceholderType.Object;
        PlaceholderShape placeholder = source.OpenXmlDocument.PresentationPart!.SlideParts.First().Slide!
            .Descendants<PlaceholderShape>().Single();
        placeholder.Type = null;

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.Equal("object", Assert.Single(conversion.Value.Slides[0].Shapes.OfType<OdpTextBox>()).PresentationClass);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "placeholder-roles" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void MasterBackgroundEffectsAreReportedInsteadOfMappedAsPlainColor() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.OpenXmlDocument.PresentationPart!.SlideMasterParts.First().SlideMaster!
            .CommonSlideData!.Background = new Background(new BackgroundProperties(
                new A.SolidFill(new A.RgbColorModelHex { Val = "336699" }), new A.EffectList()));
        source.AddSlide(PowerPointSlideLayoutType.Blank);

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.Null(conversion.Value.MasterPages[0].BackgroundColor);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
    }

    [Fact]
    public void ThemeShapeFillAndTableStyleAreExplicitLoss() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        PowerPointSlide slide = source.AddSlide(PowerPointSlideLayoutType.Blank);
        slide.AddRectanglePoints(20, 20, 100, 50);
        slide.AddTable(1, 1);
        ShapeProperties shapeProperties = source.OpenXmlDocument.PresentationPart!.SlideParts.First().Slide!
            .Descendants<ShapeProperties>().Single();
        shapeProperties.RemoveAllChildren<A.SolidFill>();
        shapeProperties.Append(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }));

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "shape-appearance" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "table-appearance" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void OdpTableCellStyleIsExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddTable(OdfRect.FromCentimeters(1, 1, 8, 3), 1, 1);
        XNamespace table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        XElement cell = source.Package.GetXml("content.xml").Descendants(table + "table-cell").Single();
        cell.SetAttributeValue(table + "style-name", "CellStyle");
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "table-appearance" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void OdpGradientShapeDoesNotProjectStaleSolidColor() {
        OdpPresentation source = OdpPresentation.Create();
        OdpRectangle rectangle = source.AddSlide().AddRectangle(OdfRect.FromCentimeters(1, 1, 5, 3));
        rectangle.FillColor = OdfColor.Parse("#336699");
        XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        XElement properties = source.Package.GetXml("content.xml")
            .Descendants(style + "graphic-properties").Single();
        properties.SetAttributeValue(draw + "fill", "gradient");
        properties.SetAttributeValue(draw + "fill-gradient-name", "Gradient1");
        source.Package.MarkXmlDirty("content.xml");

        Assert.Null(rectangle.FillColor);
        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Null(target.Slides[0].Shapes.OfType<PowerPointAutoShape>().Single().FillColor);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "shape-appearance" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void ExplicitNoFillShapeIsReportedInBothDirections() {
        using PowerPointPresentation powerPoint = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        powerPoint.AddSlide(PowerPointSlideLayoutType.Blank).AddRectanglePoints(20, 20, 100, 50);
        ShapeProperties properties = powerPoint.OpenXmlDocument.PresentationPart!.SlideParts.First().Slide!
            .Descendants<ShapeProperties>().Single();
        properties.RemoveAllChildren<A.SolidFill>();
        properties.Append(new A.NoFill());

        OdfConversionResult<OdpPresentation> fromPowerPoint = powerPoint.ToOpenDocumentResult();
        Assert.Contains(fromPowerPoint.Report.Mappings, mapping => mapping.Feature == "shape-appearance" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => powerPoint.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));

        OdpPresentation odf = OdpPresentation.Create();
        odf.AddSlide().AddRectangle(OdfRect.FromCentimeters(1, 1, 5, 3)).FillColor = null;
        OdfConversionResult<PowerPointPresentation> fromOdp = odf.ToPowerPointPresentationResult();
        using PowerPointPresentation converted = fromOdp.Value;
        Assert.Contains(fromOdp.Report.Mappings, mapping => mapping.Feature == "shape-appearance" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }
}
