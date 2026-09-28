using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using A = DocumentFormat.OpenXml.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class PowerPointOdpCurrentHeadReviewTests {
    [Fact]
    public void NonBodyNotesPlaceholderTextIsExplicitLoss() {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].Notes.Text = "Body note";
        NotesSlide notes = source.OpenXmlDocument.PresentationPart!.SlideParts.Single()
            .NotesSlidePart!.NotesSlide!;
        Shape placeholder = (Shape)notes.Descendants<Shape>().First(shape =>
            shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                .GetFirstChild<PlaceholderShape>()?.Type?.Value == PlaceholderValues.Body).CloneNode(true);
        placeholder.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<PlaceholderShape>()!.Type = PlaceholderValues.Footer;
        placeholder.TextBody!.Descendants<A.Text>().First().Text = "Authored footer";
        notes.CommonSlideData!.ShapeTree!.Append(placeholder);

        AssertPowerPointLoss(source, "notes-slide-appearance");
    }

    [Theory]
    [InlineData("fill")]
    [InlineData("outline")]
    [InlineData("geometry")]
    public void AuthoredNotesBodyPlaceholderAppearanceIsExplicitLoss(string kind) {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].Notes.Text = "Body note";
        Assert.DoesNotContain(source.ToOpenDocumentResult().Report.ForFeature("notes-slide-appearance"),
            mapping => mapping.Status == OdfConversionMappingStatus.Unsupported);
        Shape body = source.OpenXmlDocument.PresentationPart!.SlideParts.Single()
            .NotesSlidePart!.NotesSlide!.Descendants<Shape>().First(shape =>
                shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                    .GetFirstChild<PlaceholderShape>()?.Type?.Value == PlaceholderValues.Body);
        ShapeProperties properties = body.ShapeProperties ?? body.AppendChild(new ShapeProperties());
        switch (kind) {
            case "fill":
                properties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "336699" }));
                break;
            case "outline":
                properties.Append(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "336699" })));
                break;
            default:
                A.Transform2D transform = properties.GetFirstChild<A.Transform2D>()
                    ?? properties.AppendChild(new A.Transform2D());
                transform.Rotation = 5400000;
                break;
        }

        AssertPowerPointLoss(source, "notes-slide-appearance");
    }

    [Theory]
    [InlineData("bold")]
    [InlineData("alignment")]
    public void AuthoredNotesBodyTextFormattingIsExplicitLoss(string kind) {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].Notes.Text = "Body note";
        NotesSlide notes = source.OpenXmlDocument.PresentationPart!.SlideParts.Single()
            .NotesSlidePart!.NotesSlide!;
        if (kind == "bold") {
            A.Run run = notes.Descendants<A.Run>().Single();
            run.RunProperties ??= new A.RunProperties();
            run.RunProperties.Bold = true;
        } else {
            A.Paragraph paragraph = notes.Descendants<A.Paragraph>().Single();
            paragraph.ParagraphProperties ??= new A.ParagraphProperties();
            paragraph.ParagraphProperties.Alignment = A.TextAlignmentTypeValues.Center;
        }

        AssertPowerPointLoss(source, "notes-slide-appearance");
    }

    [Theory]
    [InlineData("tgtFrame", "_blank")]
    [InlineData("history", "0")]
    [InlineData("highlightClick", "1")]
    public void AuthoredRunHyperlinkFlagsAreExplicitLoss(string name, string value) {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].AddTextBoxPoints("Link", 20, 20, 100, 50);
        A.Run run = source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<A.Run>().Single();
        run.RunProperties ??= new A.RunProperties();
        A.HyperlinkOnClick click = run.RunProperties.AppendChild(new A.HyperlinkOnClick());
        click.SetAttribute(new OpenXmlAttribute("", name, "", value));

        AssertPowerPointLoss(source, "text-typography");
    }

    [Theory]
    [InlineData("prompt")]
    [InlineData("extension")]
    public void AuthoredTextPlaceholderMetadataIsExplicitLoss(string kind) {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].AddTextBoxPoints("Agenda", 20, 20, 100, 50);
        Shape shape = source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<Shape>().Single();
        PlaceholderShape placeholder = shape.NonVisualShapeProperties!
            .ApplicationNonVisualDrawingProperties!.GetFirstChild<PlaceholderShape>()
            ?? shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties
                .AppendChild(new PlaceholderShape { Type = PlaceholderValues.Body });
        if (kind == "prompt") placeholder.HasCustomPrompt = true;
        else placeholder.Append(new DocumentFormat.OpenXml.OpenXmlUnknownElement("p", "extLst",
            "http://schemas.openxmlformats.org/presentationml/2006/main"));

        AssertPowerPointLoss(source, "placeholder-metadata");
    }

    [Fact]
    public void AuthoredPresentationThumbnailIsExplicitLoss() {
        using PowerPointPresentation source = CreatePowerPoint();
        Assert.DoesNotContain(source.ToOpenDocumentResult().Report.ForFeature("presentation-thumbnail"),
            mapping => mapping.Status == OdfConversionMappingStatus.Unsupported);
        ThumbnailPart thumbnail = source.OpenXmlDocument.ThumbnailPart!;
        using (Stream stream = thumbnail.GetStream(FileMode.Create, FileAccess.Write)) {
            stream.Write(new byte[] { 0xff, 0xd8, 0xff, 0xd9 }, 0, 4);
        }

        AssertPowerPointLoss(source, "presentation-thumbnail");
    }

    [Fact]
    public void OdpTableTemplateAppearanceIsExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddTable(OdfRect.FromCentimeters(1, 1, 8, 3), 1, 1);
        XElement table = source.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table").Single();
        table.SetAttributeValue(OdfNamespaces.Table + "template-name", "BrandedTable");
        table.SetAttributeValue(OdfNamespaces.Table + "use-first-row-styles", "true");
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.ForFeature("table-appearance"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void UnfilledOdpLineDoesNotReportShapeAppearanceLoss() {
        OdpPresentation source = OdpPresentation.Create();
        OdpLine line = source.AddSlide().AddLine(OdfLength.Centimeters(1), OdfLength.Centimeters(1),
            OdfLength.Centimeters(6), OdfLength.Centimeters(2));
        line.FillColor = null;
        line.StrokeColor = OdfColor.Parse("#336699");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.DoesNotContain(conversion.Report.ForFeature("shape-appearance"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void RepeatedOdpTableRowsAndCellsConvertAtLogicalPositions() {
        OdpPresentation source = OdpPresentation.Create();
        OdpTable table = source.AddSlide().AddTable(OdfRect.FromCentimeters(1, 1, 8, 3), 1, 1);
        table.Cell(0, 0).Text = "Repeated";
        XElement tableXml = source.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table").Single();
        tableXml.Descendants(OdfNamespaces.Table + "table-row").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-rows-repeated", 2);
        tableXml.Descendants(OdfNamespaces.Table + "table-cell").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", 3);
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        PowerPointTable result = target.Slides[0].Tables.Single();
        Assert.Equal("Repeated", result.GetCell(0, 0).Paragraphs[0].Runs[0].Text);
        Assert.Equal("Repeated", result.GetCell(1, 2).Paragraphs[0].Runs[0].Text);
    }

    [Fact]
    public void RepeatedOdpTableRowsRespectLimitBeforeExpansion() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddTable(OdfRect.FromCentimeters(1, 1, 8, 3), 1, 1);
        XElement row = source.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table-row").Single();
        row.SetAttributeValue(OdfNamespaces.Table + "number-rows-repeated", 1_000);
        source.Package.MarkXmlDirty("content.xml");

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            source.ToPowerPointPresentationResult(new PowerPointOpenDocumentConversionOptions {
                MaxTableRows = 2
            }));
        Assert.Contains("configured conversion limit", exception.Message);
    }

    [Fact]
    public void PowerPointShapeHoverIsExplicitLoss() {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].AddTextBox("Hover", 0, 0, 100000, 100000);
        Shape shape = source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<Shape>().Single();
        shape.NonVisualShapeProperties!.NonVisualDrawingProperties!
            .AppendChild(new A.HyperlinkOnHover { Id = string.Empty, Action = "ppaction://hlinkshowjump?jump=nextslide" });

        AssertPowerPointLoss(source, "shape-hover-interactions");
    }

    [Fact]
    public void ActionOnlyPowerPointShapeClickIsExplicitLoss() {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].AddTextBox("Next", 0, 0, 100000, 100000);
        Shape shape = source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<Shape>().Single();
        shape.NonVisualShapeProperties!.NonVisualDrawingProperties!
            .AppendChild(new A.HyperlinkOnClick { Action = "ppaction://hlinkshowjump?jump=nextslide" });

        AssertPowerPointLoss(source, "shape-hyperlinks");
    }

    [Fact]
    public void PowerPointShapeBlackWhiteModeIsExplicitLoss() {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].AddRectanglePoints(20, 20, 100, 50);
        ShapeProperties properties = source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<Shape>().Single().ShapeProperties!;
        properties.SetAttribute(new OpenXmlAttribute("", "bwMode", "", "black"));

        AssertPowerPointLoss(source, "shape-appearance");
    }

    [Fact]
    public void PowerPointShapeUseBackgroundFillIsExplicitLoss() {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].AddRectanglePoints(20, 20, 100, 50);
        Shape shape = source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<Shape>().Single();
        shape.UseBackgroundFill = true;

        AssertPowerPointLoss(source, "shape-appearance");
    }

    [Theory]
    [InlineData(80000)]
    [InlineData(-10000)]
    public void PowerPointRunBaselineMagnitudeIsExplicitLoss(int baseline) {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].AddTextBoxPoints("Shifted", 20, 20, 100, 50);
        A.Run run = source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<A.Run>().Single();
        run.RunProperties ??= new A.RunProperties();
        run.RunProperties.Baseline = baseline;

        AssertPowerPointLoss(source, "text-typography");
    }

    [Fact]
    public void OdpTextBoxSizingConstraintsAreExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddTextBox(OdfRect.FromCentimeters(1, 1, 4, 2));
        XElement textBox = source.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Draw + "text-box").Single();
        textBox.SetAttributeValue(OdfNamespaces.Fo + "min-height", "3cm");
        source.Package.MarkXmlDirty("content.xml");

        AssertOdpLoss(source, "text-box-layout");
    }

    [Theory]
    [InlineData("rel-width", "50%")]
    [InlineData("rel-height", "75%")]
    public void RelativeOdpTextBoxFrameSizingIsExplicitLoss(string attributeName, string value) {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddTextBox(OdfRect.FromCentimeters(1, 1, 4, 2));
        XElement frame = source.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Draw + "frame").Single();
        frame.SetAttributeValue(OdfNamespaces.Style + attributeName, value);
        source.Package.MarkXmlDirty("content.xml");

        AssertOdpLoss(source, "text-box-layout");
    }

    [Fact]
    public void OdpSlideFooterDeclarationIsExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide();
        XElement page = source.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Draw + "page").Single();
        page.AddBeforeSelf(new XElement(OdfNamespaces.Presentation + "footer-decl",
            new XAttribute(OdfNamespaces.Presentation + "name", "Footer1"), "Confidential"));
        page.SetAttributeValue(OdfNamespaces.Presentation + "use-footer-name", "Footer1");
        source.Package.MarkXmlDirty("content.xml");

        AssertOdpLoss(source, "slide-headers-footers");
    }

    [Fact]
    public void InternalPowerPointTextLinkUsesPreservedOdpSlideName() {
        using PowerPointPresentation source = CreatePowerPoint();
        PowerPointSlide destination = source.AddSlide(PowerPointSlideLayoutType.Blank);
        destination.Name = "Agenda";
        PowerPointTextBox box = source.Slides[0].AddTextBoxPoints("Open agenda", 20, 20, 200, 40);
        box.Paragraphs[0].Runs[0].SetHyperlink(destination);

        OdpPresentation target = source.ToOpenDocumentResult().Value;
        XElement link = target.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Text + "a").Single();
        Assert.Equal("#Agenda", (string?)link.Attribute(OdfNamespaces.XLink + "href"));

        using PowerPointPresentation reopened = target.ToPowerPointPresentationResult().Value;
        Assert.Equal("#slide-2", reopened.Slides[0].TextBoxes.Single().Paragraphs[0].Runs[0].Hyperlink?.OriginalString);
    }

    [Fact]
    public void InternalPowerPointTextLinkUsesCollisionAdjustedOdpSlideName() {
        using PowerPointPresentation source = CreatePowerPoint();
        source.AddSlide(PowerPointSlideLayoutType.Blank).Name = "Agenda";
        PowerPointSlide destination = source.AddSlide(PowerPointSlideLayoutType.Blank);
        destination.Name = "Agenda";
        source.Slides[0].AddTextBoxPoints("Open second agenda", 20, 20, 200, 40)
            .Paragraphs[0].Runs[0].SetHyperlink(destination);

        OdpPresentation target = source.ToOpenDocumentResult().Value;
        XElement link = target.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Text + "a").Single();
        Assert.Equal("#Agenda_2", (string?)link.Attribute(OdfNamespaces.XLink + "href"));
        Assert.Equal("Agenda_2", target.Slides[2].Name);
    }

    [Fact]
    public void GroupedOdpTableRowsAreReportedBeforeFlattening() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddTable(OdfRect.FromCentimeters(1, 1, 8, 3), 1, 1).Cell(0, 0).Text = "Grouped";
        XElement row = source.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table-row").Single();
        row.ReplaceWith(new XElement(OdfNamespaces.Table + "table-row-group", row));
        source.Package.MarkXmlDirty("content.xml");

        AssertOdpLoss(source, "table-appearance");
    }

    [Fact]
    public void PowerPointTableExtensionIdentityIsExplicitLoss() {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].AddTable(1, 1);
        A.Table table = source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<A.Table>().Single();
        table.Elements<A.TableRow>().Single().AppendChild(new A.ExtensionList());

        AssertPowerPointLoss(source, "table-appearance");
    }

    [Fact]
    public void LinkedOdpTextBoxesAreExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddTextBox(OdfRect.FromCentimeters(1, 1, 4, 2)).AddParagraph("Flowing text");
        XElement textBox = source.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Draw + "text-box").Single();
        textBox.SetAttributeValue(OdfNamespaces.Draw + "chain-next-name", "NextFrame");
        source.Package.MarkXmlDirty("content.xml");

        AssertOdpLoss(source, "text-box-chains");
    }

    [Fact]
    public void OdpHeadingOutlineLevelIsExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddTextBox(OdfRect.FromCentimeters(1, 1, 4, 2)).AddParagraph("Heading");
        XElement paragraph = source.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Text + "p").Single();
        paragraph.ReplaceWith(new XElement(OdfNamespaces.Text + "h",
            new XAttribute(OdfNamespaces.Text + "outline-level", "2"), paragraph.Nodes()));
        source.Package.MarkXmlDirty("content.xml");

        AssertOdpLoss(source, "text-headings");
    }

    [Theory]
    [InlineData("text-transform", "lowercase")]
    [InlineData("text-transform", "capitalize")]
    [InlineData("text-position", "super 40%")]
    public void OdpProjectedTextEffectsAreExplicitLoss(string property, string value) {
        OdpPresentation source = OdpPresentation.Create();
        OdpTextBox box = source.AddSlide().AddTextBox(OdfRect.FromCentimeters(1, 1, 4, 2));
        OdfStyle style = source.Styles.CreateNamed("ProjectedText", OdfStyleFamily.Paragraph);
        style.Element.Add(new XElement(OdfNamespaces.Style + "text-properties",
            new XAttribute((property == "text-transform" ? OdfNamespaces.Fo : OdfNamespaces.Style) + property, value)));
        source.Package.MarkXmlDirty("styles.xml");
        OdpParagraph paragraph = box.AddParagraph("Authored text");
        paragraph.StyleName = style.Name;

        AssertOdpLoss(source, "text-effects");
    }



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

    [Theory]
    [InlineData(-0.2, 0)]
    [InlineData(2.2, 0)]
    [InlineData(1.2, 1.2)]
    [InlineData(0.9999999, 0.9999999)]
    public void UnrepresentableOdpImageCropIsExplicitLoss(double left, double right) {
        OdpPresentation source = OdpPresentation.Create();
        byte[] png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        OdpImage image = source.AddSlide().AddImage(png, "pixel.png", OdfRect.FromCentimeters(1, 1, 2, 2));
        image.Crop = new OdfInsets(OdfLength.Centimeters(0), OdfLength.Centimeters(right),
            OdfLength.Centimeters(0), OdfLength.Centimeters(left));

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.ForFeature("shape-appearance"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        if (left + right >= 1.999999D) {
            Assert.Equal(0D, Assert.Single(target.Slides[0].Pictures).CropLeftRatio);
        }
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
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

    [Theory]
    [InlineData("lang", "fr-FR")]
    [InlineData("altLang", "de-DE")]
    [InlineData("noProof", "1")]
    [InlineData("kumimoji", "1")]
    [InlineData("normalizeH", "1")]
    [InlineData("bmk", "Heading")]
    public void DirectUnsupportedRunAttributesAreExplicitLoss(string attribute, string value) {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].AddTextBoxPoints("Text", 20, 20, 100, 50);
        A.Run run = source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<A.Run>().Single();
        run.RunProperties ??= new A.RunProperties();
        run.RunProperties.SetAttribute(new OpenXmlAttribute("", attribute, "", value));

        AssertPowerPointLoss(source, "text-typography");
    }

    [Theory]
    [InlineData("kinsoku")]
    [InlineData("modifyVerifier")]
    [InlineData("smartTags")]
    public void AuthoredPresentationRootChildrenAreExplicitLoss(string childName) {
        using PowerPointPresentation source = CreatePowerPoint();
        PresentationPart presentation = source.OpenXmlDocument.PresentationPart!;
        presentation.Presentation!.AppendChild(new OpenXmlUnknownElement("p", childName,
            "http://schemas.openxmlformats.org/presentationml/2006/main"));

        AssertPowerPointLoss(source, "presentation-settings");
    }

    [Fact]
    public void AuthoredOdpHandoutMasterIsExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide();
        XDocument styles = source.Package.GetXml("styles.xml");
        XElement masterStyles = styles.Root!.Element(OdfNamespaces.Office + "master-styles")!;
        masterStyles.Add(new XElement(OdfNamespaces.Style + "handout-master",
            new XAttribute(OdfNamespaces.Style + "name", "Handout")));
        source.Package.MarkXmlDirty("styles.xml");

        AssertOdpLoss(source, "handout-master");
    }

    [Theory]
    [InlineData("pic")]
    [InlineData("graphicFrame")]
    public void NotesMediaOrGraphicalContentIsExplicitLoss(string elementName) {
        using PowerPointPresentation source = CreatePowerPoint();
        source.Slides[0].Notes.Text = "Body note";
        NotesSlide notes = source.OpenXmlDocument.PresentationPart!.SlideParts.Single().NotesSlidePart!.NotesSlide!;
        if (elementName == "pic") notes.CommonSlideData!.ShapeTree!.Append(new Picture());
        else notes.CommonSlideData!.ShapeTree!.Append(new GraphicFrame());
        AssertPowerPointLoss(source, "notes-slide-appearance");
    }

    [Theory]
    [InlineData(3)]
    [InlineData(256)]
    public void DeclaredOdpTableColumnsSurviveSparseRows(int columns) {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddTable(OdfRect.FromCentimeters(1, 1, 8, 3), 1, 1);
        source.Package.GetXml("content.xml").Descendants(OdfNamespaces.Table + "table-column").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", columns);
        source.Package.MarkXmlDirty("content.xml");
        using PowerPointPresentation converted = source.ToPowerPointPresentationResult().Value;
        Assert.Equal(columns, converted.Slides[0].Tables.Single().Columns);
        using PowerPointPresentation reopened = PowerPointPresentation.Load(new MemoryStream(converted.ToBytes()));
        Assert.Equal(columns, reopened.Slides[0].Tables.Single().Columns);
    }

    [Fact]
    public void DeclaredOdpTableColumnLimitIsEnforced() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddTable(OdfRect.FromCentimeters(1, 1, 8, 3), 1, 1);
        source.Package.GetXml("content.xml").Descendants(OdfNamespaces.Table + "table-column").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", "257");
        source.Package.MarkXmlDirty("content.xml");
        Assert.Throws<InvalidDataException>(() => source.ToPowerPointPresentationResult());
    }

    [Fact]
    public void LegacyOdpAnimationsAreInspectedAndReported() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2), "Animated");
        source.Package.GetXml("content.xml").Descendants(OdfNamespaces.Draw + "page").Single()
            .Add(new XElement(OdfNamespaces.Presentation + "animations",
                new XElement(OdfNamespaces.Presentation + "show-shape",
                    new XAttribute(OdfNamespaces.Presentation + "effect", "fade"))));
        source.Package.MarkXmlDirty("content.xml");
        Assert.Contains(source.InspectFeatures().Findings, finding => finding.Name == "presentation-animations");
        AssertOdpLoss(source, "source-presentation-animations");
    }

    [Fact]
    public void OdpCustomGluePointsAreExplicitLoss() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide().AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2), "Connection");
        source.Package.GetXml("content.xml").Descendants(OdfNamespaces.Draw + "frame").Single()
            .Add(new XElement(OdfNamespaces.Draw + "glue-point", new XAttribute(OdfNamespaces.Draw + "id", "4"),
                new XAttribute(OdfNamespaces.Svg + "x", "1cm"), new XAttribute(OdfNamespaces.Svg + "y", "1cm")));
        source.Package.MarkXmlDirty("content.xml");
        AssertOdpLoss(source, "shape-appearance");
    }

    [Fact]
    public void UnsupportedPowerPointBackgroundExplicitlyHidesOdpMasterBackground() {
        using PowerPointPresentation source = CreatePowerPoint();
        SlidePart slide = source.OpenXmlDocument.PresentationPart!.SlideParts.Single();
        slide.SlideLayoutPart!.SlideMasterPart!.SlideMaster!.CommonSlideData!.Background =
            new Background(new BackgroundProperties(new A.SolidFill(new A.RgbColorModelHex { Val = "336699" })));
        slide.Slide!.CommonSlideData!.Background = new Background(new BackgroundProperties(new A.GradientFill()));
        OdpPresentation converted = source.ToOpenDocumentResult().Value;
        OdpPresentation reopened = OdpPresentation.Load(new MemoryStream(converted.ToBytes()));
        XElement page = reopened.Package.GetXml("content.xml").Descendants(OdfNamespaces.Draw + "page").Single();
        Assert.Equal("#336699", reopened.MasterPages[0].BackgroundColor?.ToString());
        string? pageStyle = (string?)page.Attribute(OdfNamespaces.Draw + "style-name");
        XElement properties = reopened.Package.GetXml("content.xml").Descendants(OdfNamespaces.Style + "style")
            .Single(style => (string?)style.Attribute(OdfNamespaces.Style + "name") == pageStyle)
            .Element(OdfNamespaces.Style + "drawing-page-properties")!;
        Assert.Equal("false", (string?)properties.Attribute(OdfNamespaces.Presentation + "background-visible"));
        using PowerPointPresentation back = reopened.ToPowerPointPresentationResult().Value;
        Assert.Null(back.Slides[0].GetBackground().Color);
    }

    private static PowerPointPresentation CreatePowerPoint() {
        PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.AddSlide(PowerPointSlideLayoutType.Blank);
        return source;
    }

    private static void AssertOdpLoss(OdpPresentation source, string feature) {
        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == feature &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    private static void AssertPowerPointLoss(PowerPointPresentation source, string feature) {
        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == feature &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }
}
