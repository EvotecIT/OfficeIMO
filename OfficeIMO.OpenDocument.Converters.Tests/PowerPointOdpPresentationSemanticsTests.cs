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

public sealed class PowerPointOdpPresentationSemanticsTests {
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
        source.AddSlide(PowerPointSlideLayoutType.Blank)
            .AddTextBoxPoints("Inherited text", 20, 20, 200, 40);

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "masters-layouts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
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
}
