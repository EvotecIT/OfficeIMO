using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
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
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(320),OfficeIMO.OfficeLength.Points(220));
        var slide = presentation.AddSlide();
        var box = slide.AddTextBox("EDGE",global::OfficeIMO.OfficeLength.Points(x),global::OfficeIMO.OfficeLength.Points(y),global::OfficeIMO.OfficeLength.Points(width),global::OfficeIMO.OfficeLength.Points(height));
        box.FontName = "Helvetica";
        box.FontSize = 8;
        var bar = slide.AddRectanglePoints(x, y, width, height);
        var group = slide.GroupShapes(new PowerPointShape[] { bar, box });
        if (nested) {
            group.HorizontalFlip = true;
            group = slide.GroupShapes(new PowerPointShape[] { group, slide.AddRectanglePoints(x, y, width, height) });
        }
        group.Rotation = 90;
        byte[] bytes = presentation.ToPdf(new PowerPointPdfExportOptions {
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(320),OfficeIMO.OfficeLength.Points(220));
        var slide = presentation.AddSlide();
        var box = slide.AddTextBox("MAPPED",global::OfficeIMO.OfficeLength.Points(1000),global::OfficeIMO.OfficeLength.Points(1000),global::OfficeIMO.OfficeLength.Points(160),global::OfficeIMO.OfficeLength.Points(50));
        box.FontName = family;
        var group = slide.GroupShapes(new PowerPointShape[] { box, slide.AddRectanglePoints(1000, 1060, 160, 10) });
        if (nested) group = slide.GroupShapes(new PowerPointShape[] { group, slide.AddRectanglePoints(1000, 1080, 160, 10) });
        group.Left= OfficeIMO.OfficeLength.Points(40);
        group.Top= OfficeIMO.OfficeLength.Points(40);
        var result = presentation.ConvertToPdfWithEvidence(new PowerPointPdfExportOptions {
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
    public void GroupedPdfAndSnapshotRetainMatchingTextPositionAndSize(double scaleX, double scaleY, bool nested) {
        var profile = new OfficeRenderingProfile("group-text", new OfficeFontFaceCollection().Add("Fixture",
            OfficeIMO.TestAssets.PdfTestFontAssets.LoadBundledOpenTypeCffFont()));
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(600),OfficeIMO.OfficeLength.Points(500));
        PowerPointSlide slide = presentation.AddSlide();
        var box = slide.AddTextBox("INK",global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(200),global::OfficeIMO.OfficeLength.Points(60));
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
        byte[] svg = slide.ExportImage(OfficeImageExportFormat.Svg, imageOptions).Bytes;
        var snapshot = slide.CreateVisualSnapshot(imageOptions);
        var text = Assert.Single(snapshot.TextSpans, span => span.Text == "INK");
        var result = presentation.ConvertToPdfWithEvidence(new PowerPointPdfExportOptions().UseRenderingProfile(profile));
        byte[] bytes = result.ToBytes();
        var canonicalPage = OfficeIMO.Pdf.PdfReadDocument.Open(bytes).Pages[0];
        var expectedText = Assert.Single(canonicalPage.GetTextSpans(), span => span.Text == "INK");
        Assert.Equal(expectedText.X, text.X, 3);
        Assert.Equal(expectedText.Y, text.Y, 3);
        Assert.Equal(expectedText.FontSize, text.FontSize, 3);
        Assert.Equal(expectedText.Advance, text.Advance, 3);
        Assert.Equal(expectedText.RotationDegrees, text.RotationDegrees, 3);
        Assert.Equal("svg", System.Xml.Linq.XDocument.Parse(System.Text.Encoding.UTF8.GetString(svg)).Root!.Name.LocalName);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = pdf.GetPage(1);
        var letters = page.Letters;
        Assert.Equal("INK", string.Concat(letters.Select(letter => letter.Value)));
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
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(320),OfficeIMO.OfficeLength.Points(220));
        var slide = presentation.AddSlide();
        var box = slide.AddTextBox("TRANSFORM",global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(200),global::OfficeIMO.OfficeLength.Points(60));
        box.FontName = "Fixture";
        box.FontSize = 14;
        box.Paragraphs[0].Runs[0].SetHyperlink("https://officeimo.net/group");
        slide.GroupShapes(new PowerPointShape[] { box, slide.AddRectanglePoints(40, 110, 200, 10) });
        var options = new PowerPointPdfExportOptions().UseRenderingProfile(profile);
        using var before = UglyToad.PdfPig.PdfDocument.Open(presentation.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var original = before.GetPage(1).Letters[0];
        var transform = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<DocumentFormat.OpenXml.Presentation.GroupShape>().Single().GroupShapeProperties!.TransformGroup!;
        transform.Rotation = (int)(rotation * 60000D);
        transform.HorizontalFlip = horizontalFlip;
        transform.VerticalFlip = verticalFlip;
        byte[] bytes = presentation.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
