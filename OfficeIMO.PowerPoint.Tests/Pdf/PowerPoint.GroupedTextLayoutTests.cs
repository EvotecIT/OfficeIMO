using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointGroupedTextLayoutTests {
    [Theory]
    [InlineData(-20D, 60D, 100D, 40D, false)]
    [InlineData(-50D, 20D, 40D, 160D, false)]
    [InlineData(-20D, 60D, 100D, 40D, true)]
    [InlineData(-50D, 20D, 40D, 160D, true)]
    public void RotatedGroupsAtSlideEdgeRetainTextAndPageSpaceClip(double x, double y, double width, double height, bool nested) {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        presentation.SlideSize.SetSizePoints(320, 220);
        var slide = presentation.AddSlide();
        var box = slide.AddTextBoxPoints("EDGE", x, y, width, height);
        box.FontName = "Helvetica";
        box.FontSize = 8;
        var bar = slide.AddRectanglePoints(x, y, width, height);
        var group = slide.GroupShapes(new PowerPointShape[] { bar, box });
        if (nested) {
            group.HorizontalFlip = true;
            group = slide.GroupShapes(new PowerPointShape[] { group, slide.AddRectanglePoints(x, y, width, height) });
        }
        group.Rotation = 90;
        byte[] bytes = presentation.ToPdfBytes(new PowerPointToPdfOptions {
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        });
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal("EDGE", string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)));
        string operators = PdfOperatorSearchText.From(bytes);
        Assert.Contains("0 0 320 220 re W", operators, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GroupLocalFontDiscoveryReportsMappedVisibleText(bool nested) {
        const string family = "OfficeIMO Missing Group Font 81C2";
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        presentation.SlideSize.SetSizePoints(320, 220);
        var slide = presentation.AddSlide();
        var box = slide.AddTextBoxPoints("MAPPED", 1000, 1000, 160, 50);
        box.FontName = family;
        var group = slide.GroupShapes(new PowerPointShape[] { box, slide.AddRectanglePoints(1000, 1060, 160, 10) });
        if (nested) group = slide.GroupShapes(new PowerPointShape[] { group, slide.AddRectanglePoints(1000, 1080, 160, 10) });
        group.LeftPoints = 40;
        group.TopPoints = 40;
        var result = presentation.ToPdfDocumentResult(new PowerPointToPdfOptions {
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        });
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(result.ToBytes());
        Assert.Equal("MAPPED", string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)));
        Assert.Contains(result.Warnings, warning => warning.Code == "font-family-substitution"
            && warning.Details.TryGetValue("fontFamily", out string? value) && value == family);
    }

    [Theory]
    [InlineData(0.5D, 0.5D, false)]
    [InlineData(2D, 2D, false)]
    [InlineData(2D, 0.5D, false)]
    [InlineData(0.5D, 2D, true)]
    public void GroupedPdfAndSvgRetainMatchingTextPositionAndSize(double scaleX, double scaleY, bool nested) {
        var profile = new OfficeRenderingProfile("group-text", new OfficeFontFaceCollection().Add("Fixture",
            OfficeIMO.TestAssets.PdfTestFontAssets.LoadBundledOpenTypeCffFont()));
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        presentation.SlideSize.SetSizePoints(600, 500);
        PowerPointSlide slide = presentation.AddSlide();
        var box = slide.AddTextBoxPoints("INK", 40, 40, 200, 60);
        box.FontName = "Fixture";
        box.FontSize = 14;
        box.Paragraphs[0].Runs[0].Underline = true;
        var group = slide.GroupShapes(new PowerPointShape[] { box, slide.AddRectanglePoints(40, 110, 200, 10) });
        if (nested) group = slide.GroupShapes(new PowerPointShape[] { group, slide.AddRectanglePoints(40, 130, 200, 10) });
        var xmlGroup = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<DocumentFormat.OpenXml.Presentation.GroupShape>().Single();
        var transform = xmlGroup.GroupShapeProperties!.TransformGroup!;
        transform.Extents!.Cx = (long)Math.Round(transform.ChildExtents!.Cx!.Value * scaleX);
        transform.Extents.Cy = (long)Math.Round(transform.ChildExtents.Cy!.Value * scaleY);
        var imageOptions = new PowerPointImageExportOptions();
        imageOptions.UseRenderingProfile(profile);
        var svg = XDocument.Parse(Encoding.UTF8.GetString(slide.ExportImage(OfficeImageExportFormat.Svg, imageOptions).Bytes));
        XElement[] textNodes = svg.Descendants().Where(node => node.Name.LocalName == "text" && node.Value == "INK").ToArray();
        XElement text = Assert.Single(textNodes);
        var result = presentation.ToPdfDocumentResult(new PowerPointToPdfOptions().UseRenderingProfile(profile));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(result.ToBytes());
        var page = pdf.GetPage(1);
        var letters = page.Letters;
        Assert.Equal("INK", string.Concat(letters.Select(letter => letter.Value)));
        double Attribute(string name) => double.Parse(text.Attribute(name)!.Value, CultureInfo.InvariantCulture);
        Assert.InRange(Math.Abs(letters[0].StartBaseLine.X - Attribute("x")), 0D, 0.1D);
        Assert.InRange(Math.Abs(page.Height - letters[0].StartBaseLine.Y - Attribute("y")), 0D, 0.1D);
        Assert.InRange(Math.Abs(letters[0].FontSize - Attribute("font-size")), 0D, 0.1D);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "text-box-overflow");
    }

    [Theory]
    [InlineData(30D, false, false)]
    [InlineData(0D, true, false)]
    [InlineData(0D, false, true)]
    [InlineData(45D, true, true)]
    public void GroupTransformsPreserveSearchableGlyphGeometry(double rotation, bool horizontalFlip, bool verticalFlip) {
        var profile = new OfficeRenderingProfile("group-transform", new OfficeFontFaceCollection().Add("Fixture",
            OfficeIMO.TestAssets.PdfTestFontAssets.LoadBundledOpenTypeCffFont()));
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        presentation.SlideSize.SetSizePoints(320, 220);
        var slide = presentation.AddSlide();
        var box = slide.AddTextBoxPoints("TRANSFORM", 40, 40, 200, 60);
        box.FontName = "Fixture";
        box.FontSize = 14;
        box.Paragraphs[0].Runs[0].SetHyperlink("https://officeimo.net/group");
        slide.GroupShapes(new PowerPointShape[] { box, slide.AddRectanglePoints(40, 110, 200, 10) });
        var options = new PowerPointToPdfOptions().UseRenderingProfile(profile);
        using var before = UglyToad.PdfPig.PdfDocument.Open(presentation.ToPdfBytes(options));
        var original = before.GetPage(1).Letters[0];
        var transform = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<DocumentFormat.OpenXml.Presentation.GroupShape>().Single().GroupShapeProperties!.TransformGroup!;
        transform.Rotation = (int)(rotation * 60000D);
        transform.HorizontalFlip = horizontalFlip;
        transform.VerticalFlip = verticalFlip;
        byte[] bytes = presentation.ToPdfBytes(options);
        using var after = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = after.GetPage(1).Letters;
        Assert.Equal("TRANSFORM", string.Concat(letters.Select(letter => letter.Value)));
        var matrix = new OfficeImageFrameTransform(rotation, 140, 80, horizontalFlip, verticalFlip).CreateDestinationTransform();
        OfficePoint expected = matrix.TransformPoint(new OfficePoint(original.StartBaseLine.X, 220 - original.StartBaseLine.Y));
        Assert.InRange(Math.Abs(expected.X - letters[0].StartBaseLine.X), 0D, 0.1D);
        Assert.InRange(Math.Abs(expected.Y - (220 - letters[0].StartBaseLine.Y)), 0D, 0.1D);
        Assert.Equal(new[] { "https://officeimo.net/group" }, OfficeIMO.Pdf.PdfInspector.Inspect(bytes).LinkUris);
    }
}
