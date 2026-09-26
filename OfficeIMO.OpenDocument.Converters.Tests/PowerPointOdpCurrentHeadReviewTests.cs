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
