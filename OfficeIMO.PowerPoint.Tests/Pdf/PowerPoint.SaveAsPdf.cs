using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PowerPointSaveAsPdfTests {
    [Fact]
    public void SaveAsPdf_PowerPointPresentation_MapsSlideSizeTextShapeAndPictureToCanvasPdf() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 180);
        PowerPointSlide slide = presentation.AddSlide();

        PowerPointAutoShape panel = slide.AddRectanglePoints(20, 24, 120, 48);
        panel.FillColor = "EAF4FF";
        panel.OutlineColor = "1E5A96";
        panel.OutlineWidthPoints = 1.5D;

        PowerPointTextBox textBox = slide.AddTextBoxPoints("Premium Slide", 32, 36, 150, 36);
        textBox.FillColor = "FFFFFF";
        textBox.OutlineColor = "94A3B8";
        textBox.FontName = "Georgia";
        textBox.FontSize = 14;
        textBox.Color = "123456";
        textBox.Rotation = 0D;

        slide.AddPicture(new MemoryStream(CreateMinimalRgbPng()), OfficeIMO.PowerPoint.ImagePartType.Png, PowerPointUnits.FromPoints(210), PowerPointUnits.FromPoints(42), PowerPointUnits.FromPoints(50), PowerPointUnits.FromPoints(30));

        byte[] bytes = presentation.ToPdf();
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);

        Assert.Equal(1, info.PageCount);
        PdfCore.PdfPageInfo page = Assert.Single(info.Pages);
        Assert.Equal(320D, page.Width);
        Assert.Equal(180D, page.Height);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Premium Slide", text, StringComparison.Ordinal);

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("20 108 120 48 re", raw, StringComparison.Ordinal);
        AssertRawPdfContainsAnyBaseFont(raw, "Times-Roman", "Georgia");
        Assert.Contains("/Im1 Do", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_Reports_Unavailable_Font_Substitution() {
        const string unavailableFamily = "OfficeIMO Missing Font 7F0C9D";
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointTextBox textBox = presentation.AddSlide().AddTextBoxPoints("Unavailable font marker", 20, 24, 180, 40);
        textBox.FontName = unavailableFamily;

        PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult(new PowerPointPdfSaveOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost()
        });

        Assert.Contains(
            result.Warnings,
            warning => warning.Code == "font-family-substitution" &&
                       warning.Message.Contains(unavailableFamily, StringComparison.Ordinal));
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_Reports_Explicit_Font_Substitution_When_Host_Fonts_Are_Disabled() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointTextBox textBox = presentation.AddSlide().AddTextBoxPoints("Portable font marker", 20, 24, 180, 40);
        textBox.FontName = "Arial";

        PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult(new PowerPointPdfSaveOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        });

        PdfCore.PdfConversionWarning warning = Assert.Single(
            result.Warnings,
            item => item.Code == "font-family-substitution");
        Assert.Equal("Arial", warning.Details["fontFamily"]);
        Assert.Equal("Helvetica", warning.Details["fallbackSlot"]);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_Does_Not_Report_Unused_Theme_Font() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTextBox first = slide.AddTextBoxPoints("First explicit family", 20, 24, 180, 40);
        first.FontName = "Helvetica";
        PowerPointTextBox second = slide.AddTextBoxPoints("Second explicit family", 20, 74, 180, 40);
        second.FontName = "Helvetica";

        PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult(new PowerPointPdfSaveOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        });
        _ = result.ToBytes();

        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "font-family-substitution");
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesTextRunHyperlinks() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointTextBox textBox = presentation.AddSlide().AddTextBoxPoints(string.Empty, 24, 32, 150, 38);
        textBox.SetParagraphs(new[] { string.Empty });
        PowerPointTextRun run = textBox.Paragraphs[0].AddRun("OfficeIMO");
        run.SetHyperlink("https://officeimo.net/");

        PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult();
        byte[] bytes = result.ToBytes();
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);

        Assert.Equal(new[] { "https://officeimo.net/" }, info.LinkUris);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "snapshot-selective-fallback");
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_MapsTextBoxVerticalAlignmentToSharedCanvasTextBox() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(360, 200);
        PowerPointSlide slide = presentation.AddSlide();

        PowerPointTextBox top = slide.AddTextBoxPoints("TopPpt", 20, 30, 90, 90);
        top.TextVerticalAlignment = TextAnchoringTypeValues.Top;
        top.FontSize = 10;
        top.FillColor = "FFFFFF";
        top.FillTransparency = 100;

        PowerPointTextBox middle = slide.AddTextBoxPoints("MiddlePpt", 130, 30, 90, 90);
        middle.TextVerticalAlignment = TextAnchoringTypeValues.Center;
        middle.FontSize = 10;
        middle.FillColor = "FFFFFF";
        middle.FillTransparency = 100;

        PowerPointTextBox bottom = slide.AddTextBoxPoints("BottomPpt", 240, 30, 90, 90);
        bottom.TextVerticalAlignment = TextAnchoringTypeValues.Bottom;
        bottom.FontSize = 10;
        bottom.FillColor = "FFFFFF";
        bottom.FillTransparency = 100;

        byte[] bytes = presentation.ToPdf();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double topY = FindWordStartY(page, "TopPpt");
        double middleY = FindWordStartY(page, "MiddlePpt");
        double bottomY = FindWordStartY(page, "BottomPpt");

        Assert.True(topY > middleY + 30D, $"Expected PowerPoint center-anchored text to render lower than top-anchored text. Top: {topY:0.##}, middle: {middleY:0.##}.");
        Assert.True(middleY > bottomY + 30D, $"Expected PowerPoint bottom-anchored text to render lower than center-anchored text. Middle: {middleY:0.##}, bottom: {bottomY:0.##}.");
    }

    [Fact]
    public void ToPdfDocument_PowerPointPresentation_WarnsWhenParagraphTextBoxOverflows() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 140);
        PowerPointTextBox textBox = presentation.AddSlide().AddTextBoxPoints(string.Empty, 24, 28, 120, 24);
        textBox.FontSize = 14;
        textBox.FillTransparency = 100;
        var paragraphs = textBox.SetParagraphs(new[] {
            "First paragraph needs room",
            "Second paragraph should trigger overflow"
        });
        paragraphs[0].SetAlignment(TextAlignmentTypeValues.Left);
        paragraphs[1].SetAlignment(TextAlignmentTypeValues.Right);
        var options = new PowerPointPdfSaveOptions();

        PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult(options);
        result.ToBytes();

        PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings, item => item.Code == "text-box-overflow");
        Assert.Equal("Slide 1", warning.Source);
        Assert.NotNull(warning.LayoutDiagnostic);
        Assert.Equal(PdfCore.PdfLayoutDiagnosticKind.ClippedContent, warning.LayoutDiagnostic!.Kind);
        Assert.Equal("PowerPointTextBox", warning.LayoutDiagnostic.Source);
        Assert.True(warning.LayoutDiagnostic.HasBounds);
        Assert.Equal("OfficeIMO.PowerPoint.Pdf", warning.Converter);
    }

    [Fact]
    public void ToPdfDocument_PowerPointPresentation_WarnsWhenListIndentIsSimplified() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 150);
        PowerPointTextBox textBox = presentation.AddSlide().AddTextBoxPoints(string.Empty, 28, 26, 170, 76);
        textBox.FontSize = 12;
        textBox.FillTransparency = 100;
        textBox.SetBullets(
            new[] { "Indented bullet with explicit margin" },
            configure: paragraph => {
                paragraph.SetLeftMarginPoints(48);
                paragraph.SetHangingPoints(18);
            });
        var options = new PowerPointPdfSaveOptions();

        PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult(options);
        result.ToBytes();

        PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings, item => item.Code == "list-indent-simplified");
        Assert.Equal("Slide 1", warning.Source);
        Assert.NotNull(warning.LayoutDiagnostic);
        Assert.Equal(PdfCore.PdfLayoutDiagnosticKind.SimplifiedContent, warning.LayoutDiagnostic!.Kind);
        Assert.Equal("PowerPointList", warning.LayoutDiagnostic.Source);
        Assert.True(warning.LayoutDiagnostic.HasBounds);
    }

    [Fact]
    public void ToPdfDocument_PowerPointPresentation_WarnsForUnsupportedShapes() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.AddSlide().AddShape(ShapeTypeValues.Cloud, PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(50), PowerPointUnits.FromPoints(40));
        var options = new PowerPointPdfSaveOptions();

        PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult(options);
        result.ToBytes();

        PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings);
        Assert.Equal("Slide 1", warning.Source);
        Assert.Equal("unsupported-auto-shape", warning.Code);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersCommonPresetAutoShapes() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 160);
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddShapePoints(ShapeTypeValues.Triangle, 20, 24, 58, 44).Fill("1F4E79").Stroke("1F4E79", 1D);
        slide.AddShapePoints(ShapeTypeValues.Parallelogram, 96, 24, 74, 44).Fill("1976D2").Stroke("1976D2", 1D);
        slide.AddShapePoints(ShapeTypeValues.RightArrow, 36, 94, 112, 34).Fill("16A34A").Stroke("16A34A", 1D);
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.122 0.306 0.475 rg", raw, StringComparison.Ordinal);
        Assert.Contains("0.098 0.463 0.824 rg", raw, StringComparison.Ordinal);
        Assert.Contains("0.086 0.639 0.29 rg", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_ClipsPartiallyOffSlideShapes() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointAutoShape shape = presentation.AddSlide().AddRectanglePoints(-12, 20, 48, 30);
        shape.FillColor = "1E5A96";
        shape.OutlineColor = "1E5A96";
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0 110 36 30 re W", raw, StringComparison.Ordinal);
        Assert.Contains("-12 110 48 30 re", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_DegradesTinyTextBoxMarginsInsteadOfThrowing() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(160, 100);
        PowerPointTextBox textBox = presentation.AddSlide().AddTextBoxPoints("Tiny", 20, 20, 6, 6);
        textBox.FillTransparency = 100;
        var options = new PowerPointPdfSaveOptions();

        PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult(options);
        byte[] bytes = result.ToBytes();

        Assert.Contains(result.Warnings, warning => warning.Code == "text-box-padding");
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersSolidSlideBackground() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        presentation.AddSlide().BackgroundColor = "112233";
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.067 0.133 0.2 rg", raw, StringComparison.Ordinal);
        Assert.Contains("0 0 240 160 re", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersGradientSlideBackground() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        presentation.AddSlide().SetBackgroundGradient("112233", "445566", 45D);
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/SH1 sh", raw, StringComparison.Ordinal);
        Assert.Contains("/Shading", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_ResolvesThemeGradientBackgroundStops() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        presentation.SetThemeColor(PowerPointThemeColor.Accent1, "123456");
        presentation.SetThemeColor(PowerPointThemeColor.Accent2, "654321");
        PowerPointSlide slide = presentation.AddSlide();
        slide.SlidePart.Slide.CommonSlideData!.Background = new Background(
            new BackgroundProperties(
                new GradientFill(
                    new GradientStopList(
                        new GradientStop(new SchemeColor { Val = SchemeColorValues.Accent1 }) { Position = 0 },
                        new GradientStop(
                            new SchemeColor(
                                new LuminanceModulation { Val = 50000 }) { Val = SchemeColorValues.Accent2 }) { Position = 100000 }),
                    new LinearGradientFill { Angle = 5400000 })));
        slide.SlidePart.Slide.Save();
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/SH1 sh", raw, StringComparison.Ordinal);
        Assert.Contains("/Shading", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersImageSlideBackground() {
        string imagePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + ".png");
        try {
            File.WriteAllBytes(imagePath, PdfPngTestImages.CreateRgbPng(2, 1));
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 160);
            presentation.AddSlide().SetBackgroundImage(imagePath);
            var options = new PowerPointPdfSaveOptions();

            byte[] bytes = presentation.ToPdf(options);

            Assert.Empty(options.Warnings);
            string raw = Encoding.ASCII.GetString(bytes);
            Assert.Contains("/Im1 Do", raw, StringComparison.Ordinal);
            Assert.Contains("240 0 0 160 0 0 cm", raw, StringComparison.Ordinal);
        } finally {
            if (File.Exists(imagePath)) {
                File.Delete(imagePath);
            }
        }
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_StretchesCroppedImageSlideBackground() {
        string imagePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + ".png");
        try {
            File.WriteAllBytes(imagePath, PdfPngTestImages.CreateRgbPng(2, 1));
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 160);
            PowerPointSlide slide = presentation.AddSlide();
            slide.SetBackgroundImage(imagePath);
            A.BlipFill blipFill = slide.SlidePart.Slide.CommonSlideData!.Background!.BackgroundProperties!.GetFirstChild<A.BlipFill>()!;
            blipFill.SourceRectangle = new A.SourceRectangle { Left = 50000 };
            slide.SlidePart.Slide.Save();

            byte[] bytes = presentation.ToPdf();

            string raw = Encoding.ASCII.GetString(bytes);
            Assert.Contains("480 0 0 160 -240 0 cm", raw, StringComparison.Ordinal);
            Assert.Contains("0.5 0 0.5 1 re", raw, StringComparison.Ordinal);
        } finally {
            if (File.Exists(imagePath)) {
                File.Delete(imagePath);
            }
        }
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_ResolvesInheritedLayoutBackground() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointSlide slide = presentation.AddSlide();
        SlideLayoutPart layoutPart = slide.SlidePart.SlideLayoutPart!;
        layoutPart.SlideLayout.CommonSlideData ??= new CommonSlideData(new ShapeTree());
        layoutPart.SlideLayout.CommonSlideData.Background = new Background(
            new BackgroundProperties(
                new SolidFill(new RgbColorModelHex { Val = "112233" })));
        layoutPart.SlideLayout.Save();

        byte[] bytes = presentation.ToPdf();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.067 0.133 0.2 rg", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_ResolvesThemeBackgroundStyleReference() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        presentation.SetThemeColor(PowerPointThemeColor.Light1, "123456");
        PowerPointSlide slide = presentation.AddSlide();
        slide.SlidePart.Slide.CommonSlideData!.Background = new Background(
            new BackgroundStyleReference(
                new SchemeColor { Val = SchemeColorValues.Background1 }) { Index = 1001U });
        slide.SlidePart.Slide.Save();
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.071 0.204 0.337 rg", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_ResolvesDirectSchemeColorBackground() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        presentation.SetThemeColor(PowerPointThemeColor.Light2, "654321");
        PowerPointSlide slide = presentation.AddSlide();
        slide.SlidePart.Slide.CommonSlideData!.Background = new Background(
            new BackgroundProperties(
                new SolidFill(new SchemeColor { Val = SchemeColorValues.Background2 })));
        slide.SlidePart.Slide.Save();
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.396 0.263 0.129 rg", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_AppliesDirectSchemeColorBackgroundTransforms() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        presentation.SetThemeColor(PowerPointThemeColor.Light2, "654321");
        PowerPointSlide slide = presentation.AddSlide();
        slide.SlidePart.Slide.CommonSlideData!.Background = new Background(
            new BackgroundProperties(
                new SolidFill(
                    new SchemeColor(
                        new LuminanceModulation { Val = 50000 }) { Val = SchemeColorValues.Background2 })));
        slide.SlidePart.Slide.Save();
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.196 0.133 0.067 rg", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_AppliesDirectRgbBackgroundTransforms() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointSlide slide = presentation.AddSlide();
        slide.SlidePart.Slide.CommonSlideData!.Background = new Background(
            new BackgroundProperties(
                new SolidFill(
                    new RgbColorModelHex(
                        new LuminanceModulation { Val = 50000 }) { Val = "654321" })));
        slide.SlidePart.Slide.Save();
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.196 0.133 0.067 rg", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_SkipsHiddenSlidesByDefault() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointSlide visibleSlide = presentation.AddSlide();
        PowerPointAutoShape visible = visibleSlide.AddRectanglePoints(20, 24, 50, 20);
        visible.FillColor = "00AA00";
        PowerPointSlide hidden = presentation.AddSlide();
        hidden.Hidden = true;
        PowerPointAutoShape hiddenShape = hidden.AddRectanglePoints(120, 24, 50, 20);
        hiddenShape.FillColor = "FF0000";

        byte[] bytes = presentation.ToPdf();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("20 116 50 20 re", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("120 116 50 20 re", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersInheritedLayoutShapes() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointSlide slide = presentation.AddSlide();
        SlideLayoutPart layoutPart = slide.SlidePart.SlideLayoutPart!;
        ShapeTree tree = layoutPart.SlideLayout.CommonSlideData!.ShapeTree!;
        tree.AppendChild(new DocumentFormat.OpenXml.Presentation.Shape(
            new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
                new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = 700U, Name = "Layout Rule" },
                new DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new DocumentFormat.OpenXml.Presentation.ShapeProperties(
                new Transform2D(
                    new Offset { X = PowerPointUnits.FromPoints(16), Y = PowerPointUnits.FromPoints(20) },
                    new Extents { Cx = PowerPointUnits.FromPoints(50), Cy = PowerPointUnits.FromPoints(10) }),
                new PresetGeometry(new AdjustValueList()) { Preset = ShapeTypeValues.Rectangle },
                new SolidFill(new RgbColorModelHex { Val = "00AA00" }))));
        layoutPart.SlideLayout.Save();

        byte[] bytes = presentation.ToPdf();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("16 130 50 10 re", raw, StringComparison.Ordinal);
        Assert.Contains("0 0.667 0 rg", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_SkipsOverriddenInheritedPlaceholders() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointSlide slide = presentation.AddSlide();
        SlideLayoutPart layoutPart = slide.SlidePart.SlideLayoutPart!;
        ShapeTree layoutTree = layoutPart.SlideLayout.CommonSlideData!.ShapeTree!;
        layoutTree.AppendChild(new DocumentFormat.OpenXml.Presentation.Shape(
            new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
                new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = 701U, Name = "Layout Title" },
                new DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape { Type = PlaceholderValues.Title, Index = 0U })),
            new DocumentFormat.OpenXml.Presentation.ShapeProperties(
                new Transform2D(
                    new Offset { X = PowerPointUnits.FromPoints(20), Y = PowerPointUnits.FromPoints(20) },
                    new Extents { Cx = PowerPointUnits.FromPoints(160), Cy = PowerPointUnits.FromPoints(34) })),
            new DocumentFormat.OpenXml.Presentation.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph(new Run(new DocumentFormat.OpenXml.Drawing.Text("Layout Prompt"))))));
        layoutPart.SlideLayout.Save();
        PowerPointTextBox title = slide.AddTextBoxPoints("Actual Title", 20, 20, 160, 34);
        title.PlaceholderType = PlaceholderValues.Title;
        title.PlaceholderIndex = 0U;

        byte[] bytes = presentation.ToPdf();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Actual Title", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Layout Prompt", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersGroupedSlideShapes() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointAutoShape first = slide.AddRectanglePoints(20, 20, 30, 20);
        first.FillColor = "FF0000";
        PowerPointAutoShape second = slide.AddRectanglePoints(60, 20, 30, 20);
        second.FillColor = "00AA00";
        slide.GroupShapes(new PowerPointShape[] { first, second });
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("20 120 30 20 re", raw, StringComparison.Ordinal);
        Assert.Contains("60 120 30 20 re", raw, StringComparison.Ordinal);
        Assert.Contains("1 0 0 rg", raw, StringComparison.Ordinal);
        Assert.Contains("0 0.667 0 rg", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_AppliesGroupTransformToChildShapes() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointAutoShape first = slide.AddRectanglePoints(20, 20, 30, 20);
        first.FillColor = "FF0000";
        PowerPointAutoShape second = slide.AddRectanglePoints(60, 20, 30, 20);
        second.FillColor = "00AA00";
        slide.GroupShapes(new PowerPointShape[] { first, second });
        DocumentFormat.OpenXml.Presentation.GroupShape group = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<DocumentFormat.OpenXml.Presentation.GroupShape>()
            .Single();
        TransformGroup transform = group.GroupShapeProperties!.TransformGroup!;
        transform.Extents!.Cx = PowerPointUnits.FromPoints(140);
        transform.Extents.Cy = PowerPointUnits.FromPoints(40);
        transform.ChildExtents!.Cx = PowerPointUnits.FromPoints(70);
        transform.ChildExtents.Cy = PowerPointUnits.FromPoints(20);
        slide.SlidePart.Slide.Save();

        byte[] bytes = presentation.ToPdf();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("20 100 60 40 re", raw, StringComparison.Ordinal);
        Assert.Contains("100 100 60 40 re", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesInheritedLayoutTextBoxHyperlinks() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointSlide slide = presentation.AddSlide();
        SlideLayoutPart layoutPart = slide.SlidePart.SlideLayoutPart!;
        HyperlinkRelationship rel = layoutPart.AddHyperlinkRelationship(new Uri("https://officeimo.net/layout"), true);
        ShapeTree tree = layoutPart.SlideLayout.CommonSlideData!.ShapeTree!;
        tree.AppendChild(new DocumentFormat.OpenXml.Presentation.Shape(
            new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
                new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = 701U, Name = "Layout Link" },
                new DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(new ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new DocumentFormat.OpenXml.Presentation.ShapeProperties(
                new Transform2D(
                    new Offset { X = PowerPointUnits.FromPoints(24), Y = PowerPointUnits.FromPoints(32) },
                    new Extents { Cx = PowerPointUnits.FromPoints(150), Cy = PowerPointUnits.FromPoints(36) }),
                new PresetGeometry(new AdjustValueList()) { Preset = ShapeTypeValues.Rectangle }),
            new DocumentFormat.OpenXml.Presentation.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph(
                    new Run(
                        new RunProperties(new HyperlinkOnClick { Id = rel.Id }),
                        new DocumentFormat.OpenXml.Drawing.Text("Layout Link"))))));
        layoutPart.SlideLayout.Save();

        byte[] bytes = presentation.ToPdf();
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);

        Assert.Equal(new[] { "https://officeimo.net/layout" }, info.LinkUris);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersInheritedLayoutPresetAutoShapes() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointSlide slide = presentation.AddSlide();
        SlideLayoutPart layoutPart = slide.SlidePart.SlideLayoutPart!;
        ShapeTree tree = layoutPart.SlideLayout.CommonSlideData!.ShapeTree!;
        tree.AppendChild(new DocumentFormat.OpenXml.Presentation.Shape(
            new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
                new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = 702U, Name = "Layout Triangle" },
                new DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new DocumentFormat.OpenXml.Presentation.ShapeProperties(
                new Transform2D(
                    new Offset { X = PowerPointUnits.FromPoints(30), Y = PowerPointUnits.FromPoints(44) },
                    new Extents { Cx = PowerPointUnits.FromPoints(72), Cy = PowerPointUnits.FromPoints(48) }),
                new PresetGeometry(new AdjustValueList()) { Preset = ShapeTypeValues.Triangle },
                new SolidFill(new RgbColorModelHex { Val = "1F4E79" }))));
        layoutPart.SlideLayout.Save();
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.122 0.306 0.475 rg", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersTextBearingAutoShapeGeometry() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTextBox textBox = slide.AddTextBoxPoints("Rounded Label", 30, 40, 100, 36);
        textBox.FillColor = "FDE68A";
        textBox.OutlineColor = "92400E";
        textBox.FontSize = 12;
        var shape = (DocumentFormat.OpenXml.Presentation.Shape)textBox.Element;
        shape.ShapeProperties!.GetFirstChild<PresetGeometry>()!.Preset = ShapeTypeValues.RoundRectangle;
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Contains("Rounded Label", pdf.GetPage(1).Text, StringComparison.Ordinal);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains(" c", raw, StringComparison.Ordinal);
        Assert.Contains("0.992 0.902 0.541 rg", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_HonorsDisabledInheritedLayoutShapes() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointSlide slide = presentation.AddSlide();
        slide.SlidePart.Slide.ShowMasterShapes = false;
        SlideLayoutPart layoutPart = slide.SlidePart.SlideLayoutPart!;
        ShapeTree tree = layoutPart.SlideLayout.CommonSlideData!.ShapeTree!;
        tree.AppendChild(new DocumentFormat.OpenXml.Presentation.Shape(
            new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
                new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = 701U, Name = "Hidden Layout Rule" },
                new DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new DocumentFormat.OpenXml.Presentation.ShapeProperties(
                new Transform2D(
                    new Offset { X = PowerPointUnits.FromPoints(16), Y = PowerPointUnits.FromPoints(20) },
                    new Extents { Cx = PowerPointUnits.FromPoints(50), Cy = PowerPointUnits.FromPoints(10) }),
                new PresetGeometry(new AdjustValueList()) { Preset = ShapeTypeValues.Rectangle },
                new SolidFill(new RgbColorModelHex { Val = "00AA00" }))));
        layoutPart.SlideLayout.Save();

        byte[] bytes = presentation.ToPdf();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.DoesNotContain("16 130 50 10 re", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("0 0.667 0 rg", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_ResolvesSlidePlaceholderBoundsFromLayout() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        PowerPointSlide slide = presentation.AddSlide();
        slide.SetLayout(SlideLayoutValues.Text);
        slide.AddTextBoxPoints("Layout Bound Title", 12, 12, 120, 30);
        DocumentFormat.OpenXml.Presentation.Shape placeholderShape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<DocumentFormat.OpenXml.Presentation.Shape>()
            .Last(shape => shape.TextBody?.InnerText.Contains("Layout Bound Title", StringComparison.Ordinal) == true);
        placeholderShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties ??= new ApplicationNonVisualDrawingProperties();
        placeholderShape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.PlaceholderShape =
            new PlaceholderShape { Type = PlaceholderValues.Title };
        placeholderShape.ShapeProperties!.Transform2D?.Remove();

        var options = new PowerPointPdfSaveOptions();
        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Contains("Layout Bound Title", pdf.GetPage(1).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesFlippedPictures() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointPicture picture = presentation.AddSlide().AddPicture(
            new MemoryStream(CreateMinimalRgbPng()),
            OfficeIMO.PowerPoint.ImagePartType.Png,
            PowerPointUnits.FromPoints(40),
            PowerPointUnits.FromPoints(50),
            PowerPointUnits.FromPoints(60),
            PowerPointUnits.FromPoints(30));
        picture.HorizontalFlip = true;
        var options = new PowerPointPdfSaveOptions {
            PictureFit = OfficeImageFit.Stretch,
            WarnOnPictureAspectRatioDistortion = false
        };

        byte[] bytes = presentation.ToPdf(options);

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("-60 0 0 30 100 80 cm", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_ExplicitContainPictureFitPreservesAspectRatio() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(200, 160);
        presentation.AddSlide().AddPicture(
            new MemoryStream(PdfPngTestImages.CreateRgbPng(2, 1)),
            OfficeIMO.PowerPoint.ImagePartType.Png,
            PowerPointUnits.FromPoints(40),
            PowerPointUnits.FromPoints(40),
            PowerPointUnits.FromPoints(80),
            PowerPointUnits.FromPoints(80));
        var options = new PowerPointPdfSaveOptions {
            PictureFit = OfficeImageFit.Contain
        };

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("80 0 0 40 40 60 cm", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_DefaultPictureFitMatchesAuthoredFrame() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(200, 160);
        presentation.AddSlide().AddPicture(
            new MemoryStream(PdfPngTestImages.CreateRgbPng(2, 1)),
            OfficeIMO.PowerPoint.ImagePartType.Png,
            PowerPointUnits.FromPoints(40),
            PowerPointUnits.FromPoints(40),
            PowerPointUnits.FromPoints(80),
            PowerPointUnits.FromPoints(80));
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("80 0 0 80 40 40 cm", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_WarnsWhenExplicitStretchDistortsPictureAspectRatio() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(200, 160);
        presentation.AddSlide().AddPicture(
            new MemoryStream(PdfPngTestImages.CreateRgbPng(2, 1)),
            OfficeIMO.PowerPoint.ImagePartType.Png,
            PowerPointUnits.FromPoints(40),
            PowerPointUnits.FromPoints(40),
            PowerPointUnits.FromPoints(80),
            PowerPointUnits.FromPoints(80));
        var options = new PowerPointPdfSaveOptions {
            PictureFit = OfficeImageFit.Stretch,
            WarnOnPictureAspectRatioDistortion = true
        };

        PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult(options);
        byte[] bytes = result.ToBytes();

        PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings, item => item.Code == "picture-aspect-distortion");
        Assert.Equal("Slide 1", warning.Source);
        Assert.NotNull(warning.LayoutDiagnostic);
        Assert.Equal(PdfCore.PdfLayoutDiagnosticKind.SimplifiedContent, warning.LayoutDiagnostic!.Kind);
        Assert.Equal("PowerPointPicture", warning.LayoutDiagnostic.Source);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("80 0 0 80 40 40 cm", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesPictureHyperlinks() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddPicture(
            new MemoryStream(CreateMinimalRgbPng()),
            OfficeIMO.PowerPoint.ImagePartType.Png,
            PowerPointUnits.FromPoints(30),
            PowerPointUnits.FromPoints(40),
            PowerPointUnits.FromPoints(50),
            PowerPointUnits.FromPoints(30));
        HyperlinkRelationship rel = slide.SlidePart.AddHyperlinkRelationship(new Uri("https://officeimo.net/picture"), true);
        DocumentFormat.OpenXml.Presentation.Picture picture = slide.SlidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Single();
        picture.NonVisualPictureProperties!.NonVisualDrawingProperties!.Append(new HyperlinkOnClick { Id = rel.Id });
        slide.SlidePart.Slide.Save();

        byte[] bytes = presentation.ToPdf();
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);

        PdfCore.PdfLinkAnnotation link = Assert.Single(info.LinkAnnotations);
        Assert.Equal("https://officeimo.net/picture", link.Uri);
        Assert.Equal(30D, link.X1, 1);
        Assert.Equal(90D, link.Y1, 1);
        Assert.Equal(80D, link.X2, 1);
        Assert.Equal(120D, link.Y2, 1);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersPictureWithAltTextWithoutHyperlink() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointPicture picture = presentation.AddSlide().AddPicture(
            new MemoryStream(CreateMinimalRgbPng()),
            OfficeIMO.PowerPoint.ImagePartType.Png,
            PowerPointUnits.FromPoints(30),
            PowerPointUnits.FromPoints(40),
            PowerPointUnits.FromPoints(50),
            PowerPointUnits.FromPoints(30));
        picture.AltText = "Logo alt";
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Im1 Do", raw, StringComparison.Ordinal);
        Assert.Contains("/Figure << /Alt <4C6F676F20616C74> >> BDC", raw, StringComparison.Ordinal);
        Assert.Empty(PdfCore.PdfInspector.Inspect(bytes).LinkAnnotations);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_AppliesPictureCrop() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointPicture picture = presentation.AddSlide().AddPicture(
            new MemoryStream(CreateMinimalRgbPng()),
            OfficeIMO.PowerPoint.ImagePartType.Png,
            PowerPointUnits.FromPoints(40),
            PowerPointUnits.FromPoints(50),
            PowerPointUnits.FromPoints(60),
            PowerPointUnits.FromPoints(30));
        picture.Crop(leftPercent: 50D, topPercent: 0D, rightPercent: 0D, bottomPercent: 0D);

        byte[] bytes = presentation.ToPdf();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("120 0 0 30 -20 80 cm", raw, StringComparison.Ordinal);
        Assert.Contains("0.5 0 0.5 1 re", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RotatesPictureCropWithImageFrame() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointPicture picture = presentation.AddSlide().AddPicture(
            new MemoryStream(CreateMinimalRgbPng()),
            OfficeIMO.PowerPoint.ImagePartType.Png,
            PowerPointUnits.FromPoints(40),
            PowerPointUnits.FromPoints(50),
            PowerPointUnits.FromPoints(60),
            PowerPointUnits.FromPoints(30));
        picture.Crop(leftPercent: 25D, topPercent: 0D, rightPercent: 0D, bottomPercent: 0D);
        picture.Rotation = 90D;

        byte[] bytes = presentation.ToPdf();

        string raw = Encoding.ASCII.GetString(bytes);
        int imageTransform = raw.IndexOf("0 80 -30 0 75 55 cm", StringComparison.Ordinal);
        int localClip = raw.IndexOf("0.25 0 0.75 1 re", StringComparison.Ordinal);

        Assert.True(imageTransform >= 0, "Expected the cropped picture to render through the rotated image transform.");
        Assert.True(localClip > imageTransform, "Expected the source crop clip to be applied inside the rotated image frame.");
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesParagraphAlignmentAndListMarkers() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 180);
        PowerPointTextBox textBox = presentation.AddSlide().AddTextBoxPoints(string.Empty, 30, 36, 170, 70);
        textBox.FillTransparency = 100;
        textBox.OutlineColor = null;
        textBox.FontSize = 10;
        textBox.SetParagraphs(new[] { "Heading", "Item" });
        textBox.Paragraphs[0].Alignment = TextAlignmentTypeValues.Center;
        textBox.Paragraphs[1].Alignment = TextAlignmentTypeValues.Left;
        textBox.Paragraphs[1].SetBullet('*');

        byte[] bytes = presentation.ToPdf();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        string text = string.Join("", page.Letters.Select(letter => letter.Value));
        Assert.Contains("* Item", text, StringComparison.Ordinal);
        Assert.True(FindWordStartX(page, "Heading") > FindWordStartX(page, "Item") + 20D, "Expected centered heading text to start to the right of the left-aligned bullet item.");
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesTextBoxLineBreaks() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 180);
        PowerPointTextBox textBox = presentation.AddSlide().AddTextBoxPoints(string.Empty, 30, 32, 160, 96);
        textBox.FillTransparency = 100;
        textBox.OutlineColor = null;
        textBox.FontSize = 12;
        textBox.SetParagraphs(new[] { string.Empty });
        DocumentFormat.OpenXml.Drawing.Paragraph paragraph = textBox.Paragraphs[0].Paragraph;
        paragraph.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.Run>();
        paragraph.Append(
            new DocumentFormat.OpenXml.Drawing.Run(new DocumentFormat.OpenXml.Drawing.Text("First")),
            new DocumentFormat.OpenXml.Drawing.Break(),
            new DocumentFormat.OpenXml.Drawing.Run(new DocumentFormat.OpenXml.Drawing.Text("Second")));

        byte[] bytes = presentation.ToPdf();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double firstY = FindWordStartY(page, "First");
        double secondY = FindWordStartY(page, "Second");

        Assert.True(firstY > secondY, "Expected an explicit PowerPoint text box line break to render the following run on a lower line.");
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesTextBoxFillTransparency() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointTextBox textBox = presentation.AddSlide().AddTextBoxPoints("Transparent", 20, 30, 120, 40);
        textBox.FillColor = "112233";
        textBox.FillTransparency = 50;
        textBox.OutlineColor = null;

        byte[] bytes = presentation.ToPdf();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/ca 0.5", raw, StringComparison.Ordinal);
        Assert.Contains("0.067 0.133 0.2 rg", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesAsymmetricTextBoxMargins() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 180);
        PowerPointTextBox textBox = presentation.AddSlide().AddTextBoxPoints("Asymmetric", 30, 40, 140, 50);
        textBox.FillTransparency = 100;
        textBox.OutlineColor = null;
        textBox.FontSize = 10;
        textBox.SetTextMarginsPoints(left: 20D, top: 6D, right: 4D, bottom: 2D);

        byte[] bytes = presentation.ToPdf();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("50 92 116 42 re", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_PowerPointPresentation_Uses_Theme_Default_Text_Typography() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 180);
        PowerPointTextBox textBox = presentation.AddSlide().AddTextBoxPoints("Theme default", 30, 34, 180, 46);
        textBox.FillTransparency = 100;
        textBox.OutlineColor = null;

        PdfCore.PdfDocument pdfDocument = presentation.ToPdfDocument();

        var canvas = Assert.IsType<PdfCore.PdfCanvasBlock>(Assert.Single(pdfDocument.Blocks));
        PdfCore.PdfCanvasTextBoxItem textItem = Assert.Single(canvas.Items.OfType<PdfCore.PdfCanvasTextBoxItem>());
        Assert.Equal(18D, textItem.Style.FontSize);

        string expectedThemeFont = presentation.GetThemeLatinFonts().MinorLatin!;
        Assert.False(string.IsNullOrWhiteSpace(expectedThemeFont));
        Assert.All(textItem.Runs, run => Assert.Equal(expectedThemeFont, run.FontFamily));

        byte[] bytes = pdfDocument.ToBytes();
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.InRange(AverageLetterFontSize(pdf.GetPage(1), "Theme"), 17.5D, 18.5D);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersTablesThroughSharedPdfCanvasTable() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 180);
        PowerPointTable table = presentation.AddSlide().AddTablePoints(2, 2, 30, 34, 150, 70);
        table.FirstRow = true;
        table.BandedRows = false;
        table.SetColumnWidthsPoints(90, 60);
        table.SetRowHeightsPoints(28, 42);

        PowerPointTableCell header = table.GetCell(0, 0);
        header.Text = "Metric";
        header.FillColor = "D9EAF7";
        header.Bold = true;

        PowerPointTableCell headerScore = table.GetCell(0, 1);
        headerScore.Text = "Score";
        headerScore.FillColor = "D9EAF7";
        headerScore.HorizontalAlignment = TextAlignmentTypeValues.Center;

        PowerPointTableCell body = table.GetCell(1, 0);
        body.Text = "Quality";
        body.PaddingLeftPoints = 8D;
        body.BorderColor = "1E5A96";

        PowerPointTableCell score = table.GetCell(1, 1);
        score.Text = "99";
        score.FillColor = "EAF4FF";
        score.HorizontalAlignment = TextAlignmentTypeValues.Center;
        score.VerticalAlignment = TextAnchoringTypeValues.Center;

        var options = new PowerPointPdfSaveOptions();
        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("30 76 150 70 re", raw, StringComparison.Ordinal);
        Assert.Contains("120 146 m", raw, StringComparison.Ordinal);
        Assert.Contains("120 76 l", raw, StringComparison.Ordinal);
        Assert.Contains("30 118 m", raw, StringComparison.Ordinal);
        Assert.Contains("180 118 l", raw, StringComparison.Ordinal);
        Assert.Contains("120 76 60 42 re", raw, StringComparison.Ordinal);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Metric", text, StringComparison.Ordinal);
        Assert.Contains("Quality", text, StringComparison.Ordinal);
        Assert.Contains("99", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_PowerPointPresentation_Uses_Theme_Table_Style_And_Default_Typography() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 180);
        PowerPointTable table = presentation.AddSlide().AddTablePoints(2, 2, 30, 34, 150, 70);
        table.FirstRow = true;
        table.BandedRows = true;
        table.GetCell(0, 0).Text = "Header";
        table.GetCell(0, 1).Text = "State";
        table.GetCell(1, 0).Text = "Body";
        table.GetCell(1, 1).Text = "Ready";

        PdfCore.PdfDocument pdfDocument = presentation.ToPdfDocument();

        var canvas = Assert.IsType<PdfCore.PdfCanvasBlock>(Assert.Single(pdfDocument.Blocks));
        PdfCore.PdfCanvasTableItem tableItem = Assert.Single(canvas.Items.OfType<PdfCore.PdfCanvasTableItem>());
        PdfCore.PdfTableStyle style = Assert.IsType<PdfCore.PdfTableStyle>(tableItem.Block.Style);
        Assert.Equal(18D, style.FontSize);
        Assert.Equal(18D, style.HeaderFontSize);

        Assert.NotNull(style.CellFills);
        PdfCore.PdfColor headerFill = style.CellFills![(0, 0)];
        PdfCore.PdfColor bodyFill = style.CellFills[(1, 0)];
        Assert.NotEqual(headerFill, bodyFill);
        Assert.NotEqual(PdfCore.PdfColor.White, headerFill);

        Assert.NotNull(style.CellBorders);
        PdfCore.PdfCellBorder headerBorder = style.CellBorders![(0, 0)];
        Assert.NotNull(headerBorder.BottomBorder);
        Assert.True(headerBorder.BottomBorder!.Width >= 1D);

        string expectedThemeFont = presentation.GetThemeLatinFonts().MinorLatin!;
        Assert.False(string.IsNullOrWhiteSpace(expectedThemeFont));
        foreach (PdfCore.PdfTableCell cell in tableItem.Block.Cells.SelectMany(row => row)) {
            Assert.All(cell.Runs, run => Assert.Equal(expectedThemeFont, run.FontFamily));
        }

        byte[] bytes = pdfDocument.ToBytes();
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        double headerSize = AverageLetterFontSize(pdf.GetPage(1), "Header");
        double bodySize = AverageLetterFontSize(pdf.GetPage(1), "Body");
        Assert.InRange(headerSize, 17.5D, 18.5D);
        Assert.InRange(bodySize, 17.5D, 18.5D);
    }

    [Fact]
    public void ToPdfDocument_PowerPointPresentation_Uses_Configured_Default_Table_Style() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 180);
        PowerPointTable table = presentation.AddSlide().AddTablePoints(2, 2, 30, 34, 150, 70);
        table.FirstRow = true;
        table.BandedRows = false;
        table.SetColumnWidthsPoints(90, 60);
        table.SetRowHeightsPoints(28, 42);
        table.GetCell(0, 0).Text = "Metric";
        table.GetCell(0, 1).Text = "Score";
        table.GetCell(1, 0).Text = "Quality";
        table.GetCell(1, 1).Text = "99";

        var configuredStyle = new PdfCore.PdfTableStyle {
            CellPaddingX = 8D,
            CellPaddingY = 6D,
            BorderColor = null,
            HeaderFill = PdfCore.PdfColor.FromRgb(10, 20, 30),
            HeaderTextColor = PdfCore.PdfColor.FromRgb(240, 245, 250),
            RowStripeFill = PdfCore.PdfColor.FromRgb(220, 235, 250),
            CellFills = new Dictionary<(int Row, int Column), PdfCore.PdfColor> {
                [(1, 1)] = PdfCore.PdfColor.FromRgb(200, 210, 220)
            },
            FontSize = 12.5D,
            SpacingAfter = 11D
        };

        PdfCore.PdfDocument pdfDocument = presentation.ToPdfDocument(new PowerPointPdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions {
                DefaultTableStyle = configuredStyle
            }
        });

        var canvas = Assert.IsType<PdfCore.PdfCanvasBlock>(Assert.Single(pdfDocument.Blocks));
        PdfCore.PdfCanvasTableItem tableItem = Assert.Single(canvas.Items.OfType<PdfCore.PdfCanvasTableItem>());
        Assert.NotNull(tableItem.Block.Style);
        PdfCore.PdfTableStyle style = tableItem.Block.Style!;

        Assert.Equal(1, style.HeaderRowCount);
        Assert.Equal(8D, style.CellPaddingX);
        Assert.Equal(6D, style.CellPaddingY);
        Assert.Null(style.BorderColor);
        Assert.Equal(PdfCore.PdfColor.FromRgb(10, 20, 30), style.HeaderFill);
        Assert.Equal(PdfCore.PdfColor.FromRgb(240, 245, 250), style.HeaderTextColor);
        Assert.Null(style.RowStripeFill);
        Assert.Equal(12.5D, style.FontSize);
        Assert.Equal(11D, style.SpacingAfter);
        Assert.Equal(new double?[] { 90D, 60D }, style.ColumnWidthPoints);
        Assert.Equal(new double?[] { 28D, 42D }, style.RowMinHeights);

        Assert.Equal(PdfCore.PdfColor.FromRgb(220, 235, 250), configuredStyle.RowStripeFill);
        Assert.Null(configuredStyle.ColumnWidthPoints);
        Assert.Null(configuredStyle.RowMinHeights);
        Assert.NotNull(style.CellFills);
        Assert.Equal(PdfCore.PdfColor.FromRgb(200, 210, 220), style.CellFills![(1, 1)]);
        Assert.Equal(PdfCore.PdfColor.FromRgb(200, 210, 220), configuredStyle.CellFills![(1, 1)]);
    }

    [Fact]
    public void ToPdfDocument_PowerPointPresentation_Resets_Repeating_Header_Count_From_Slide_Table() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 180);
        PowerPointTable table = presentation.AddSlide().AddTablePoints(1, 1, 30, 34, 150, 40);
        table.FirstRow = false;
        table.GetCell(0, 0).Text = "Value";

        PdfCore.PdfDocument pdfDocument = presentation.ToPdfDocument(new PowerPointPdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions {
                DefaultTableStyle = new PdfCore.PdfTableStyle {
                    HeaderRowCount = 2,
                    RepeatHeaderRowCount = 2
                }
            }
        });

        var canvas = Assert.IsType<PdfCore.PdfCanvasBlock>(Assert.Single(pdfDocument.Blocks));
        PdfCore.PdfTableStyle style = Assert.Single(canvas.Items.OfType<PdfCore.PdfCanvasTableItem>()).Block.Style!;

        Assert.Equal(0, style.HeaderRowCount);
        Assert.Equal(0, style.RepeatHeaderRowCount);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_SkipsExcludedShapeFontsBeforeRendering() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 180);
        PowerPointSlide slide = presentation.AddSlide();

        PowerPointTextBox excluded = slide.AddTextBoxPoints("Excluded", 30, 26, 100, 28);
        excluded.FontName = "Georgia";

        PowerPointTable table = slide.AddTablePoints(1, 1, 30, 74, 150, 42);
        PowerPointTableCell cell = table.GetCell(0, 0);
        cell.Text = "Visible";
        cell.FontName = "Times New Roman";

        byte[] bytes = presentation.ToPdf(new PowerPointPdfSaveOptions {
            IncludeTextBoxes = false,
            IncludeTables = true
        });

        string raw = Encoding.ASCII.GetString(bytes);
        AssertRawPdfContainsAnyBaseFont(raw, "Times");
        Assert.DoesNotContain("Georgia", raw, StringComparison.OrdinalIgnoreCase);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Visible", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Excluded", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_SkipsUnrenderableShapeFontsBeforeRendering() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 180);
        PowerPointSlide slide = presentation.AddSlide();

        PowerPointTextBox offSlide = slide.AddTextBoxPoints("OffSlide", -180, 24, 80, 28);
        offSlide.FontName = "Georgia";

        PowerPointTable table = slide.AddTablePoints(1, 1, 30, 74, 150, 42);
        PowerPointTableCell cell = table.GetCell(0, 0);
        cell.Text = "Visible";
        cell.FontName = "Times New Roman";

        byte[] bytes = presentation.ToPdf(new PowerPointPdfSaveOptions {
            IncludeTextBoxes = true,
            IncludeTables = true
        });

        string raw = Encoding.ASCII.GetString(bytes);
        AssertRawPdfContainsAnyBaseFont(raw, "Times");
        Assert.DoesNotContain("Georgia", raw, StringComparison.OrdinalIgnoreCase);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Visible", text, StringComparison.Ordinal);
        Assert.DoesNotContain("OffSlide", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesConfiguredDefaultFontSlot() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 180);
        PowerPointSlide slide = presentation.AddSlide();

        PowerPointTextBox styled = slide.AddTextBoxPoints("StyledSerif", 30, 34, 120, 28);
        styled.FontName = "Georgia";
        PowerPointTextBox plain = slide.AddTextBoxPoints("DefaultSerif", 30, 84, 120, 28);

        byte[] bytes = presentation.ToPdf(new PowerPointPdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions {
                DefaultFont = PdfCore.PdfStandardFont.TimesRoman
            }
                .RegisterNamedFontFamily(new PdfCore.PdfEmbeddedFontFamily(
                    "Georgia",
                    OfficeIMO.TestAssets.PdfTestFontAssets.LoadBundledOpenTypeCffFont()))
        });

        string raw = Encoding.ASCII.GetString(bytes);
        AssertRawPdfContainsAnyBaseFont(raw, "Times");
        AssertRawPdfContainsAnyBaseFont(raw, "Georgia");

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("StyledSerif", text, StringComparison.Ordinal);
        Assert.Contains("DefaultSerif", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesExplicitMappedDefaultFontFamily() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 180);
        presentation.AddSlide().AddTextBoxPoints("ExplicitSerif", 30, 40, 150, 36);

        byte[] bytes = presentation.ToPdf(new PowerPointPdfSaveOptions {
            FontFamily = "serif"
        });

        string raw = Encoding.ASCII.GetString(bytes);
        AssertRawPdfContainsAnyBaseFont(raw, "Times");
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_UsesSansFallbackForUnmappedExplicitFonts() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 180);
        PowerPointTextBox textBox = presentation.AddSlide().AddTextBoxPoints("VisibleSans", 30, 40, 150, 36);
        textBox.FontName = "Aptos Display";

        byte[] bytes = presentation.ToPdf(new PowerPointPdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions {
                DefaultFont = PdfCore.PdfStandardFont.TimesRoman
            }
        });

        string raw = Encoding.ASCII.GetString(bytes);
        AssertRawPdfContainsAnyBaseFont(raw, "Helvetica");

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("VisibleSans", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesTableCellLineBreaks() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 180);
        PowerPointTable table = presentation.AddSlide().AddTablePoints(1, 1, 30, 28, 150, 96);
        table.SetRowHeightsPoints(96);

        PowerPointTableCell cell = table.GetCell(0, 0);
        cell.FontSize = 12;
        cell.Text = "First";
        DocumentFormat.OpenXml.Drawing.TextBody textBody = cell.Cell.TextBody!;
        textBody.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.Paragraph>();
        textBody.Append(new DocumentFormat.OpenXml.Drawing.Paragraph(
            new DocumentFormat.OpenXml.Drawing.Run(new DocumentFormat.OpenXml.Drawing.Text("First")),
            new DocumentFormat.OpenXml.Drawing.Break(),
            new DocumentFormat.OpenXml.Drawing.Run(new DocumentFormat.OpenXml.Drawing.Text("Second"))));
        textBody.Append(new DocumentFormat.OpenXml.Drawing.Paragraph(
            new DocumentFormat.OpenXml.Drawing.Run(new DocumentFormat.OpenXml.Drawing.Text("Third"))));

        byte[] bytes = presentation.ToPdf();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double firstY = FindWordStartY(page, "First");
        double secondY = FindWordStartY(page, "Second");
        double thirdY = FindWordStartY(page, "Third");

        Assert.True(firstY > secondY, "Expected an explicit PowerPoint table cell line break to render the following run on a lower line.");
        Assert.True(secondY > thirdY, "Expected a second PowerPoint table cell paragraph to render on a lower line.");
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesPowerPointTableCellRunFormatting() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 180);
        PowerPointTable table = presentation.AddSlide().AddTablePoints(1, 1, 30, 34, 180, 70);
        PowerPointTableCell cell = table.GetCell(0, 0);
        cell.Text = "SmallLarge";
        cell.FontSize = 10;
        A.TextBody textBody = cell.Cell.TextBody!;
        textBody.RemoveAllChildren<A.Paragraph>();
        textBody.Append(new A.Paragraph(
            new A.Run(
                new A.RunProperties { FontSize = 1000 },
                new A.Text("Small")),
            new A.Run(
                new A.RunProperties { FontSize = 1800, Bold = true },
                new A.Text("Large"))));

        byte[] bytes = presentation.ToPdf();
        string raw = Encoding.ASCII.GetString(bytes);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double smallSize = AverageLetterFontSize(page, "Small");
        double largeSize = AverageLetterFontSize(page, "Large");
        Assert.True(largeSize > smallSize + 4D, $"Expected table rich run font size to flow into PDF output. Small: {smallSize:0.##}, large: {largeSize:0.##}.");
    }

    [Fact]
    public void ToPdfDocument_PowerPointPresentation_WarnsWhenTableCellTextMayOverflow() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 150);
        PowerPointTable table = presentation.AddSlide().AddTablePoints(1, 1, 30, 34, 86, 22);
        PowerPointTableCell cell = table.GetCell(0, 0);
        cell.Text = "A very dense PowerPoint table cell that cannot fit inside this tiny fixed PDF frame";
        cell.FontSize = 14;
        var options = new PowerPointPdfSaveOptions();

        PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult(options);
        result.ToBytes();

        PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings, item => item.Code == "table-cell-overflow");
        Assert.Equal("Slide 1", warning.Source);
        Assert.NotNull(warning.LayoutDiagnostic);
        Assert.Equal(PdfCore.PdfLayoutDiagnosticKind.ClippedContent, warning.LayoutDiagnostic!.Kind);
        Assert.Equal("PowerPointTableCell", warning.LayoutDiagnostic.Source);
        Assert.True(warning.LayoutDiagnostic.HasBounds);
        Assert.Equal("OfficeIMO.PowerPoint.Pdf", warning.Converter);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesTableRotation() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 180);
        PowerPointTable table = presentation.AddSlide().AddTablePoints(1, 1, 30, 34, 150, 70);
        table.Rotation = 90D;
        table.GetCell(0, 0).Text = "Rotated";

        byte[] bytes = presentation.ToPdf();

        string raw = Encoding.ASCII.GetString(bytes);
        int transform = raw.IndexOf("0 1 -1 0 216 6 cm", StringComparison.Ordinal);
        int tableRect = raw.IndexOf("30 76 150 70 re", StringComparison.Ordinal);

        Assert.True(transform >= 0, "Expected PowerPoint table rotation to flow into the shared PDF canvas table.");
        Assert.True(tableRect > transform, "Expected rotated table geometry to render after the rotation matrix.");
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersChartsThroughSharedVectorRenderer() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 240);
        var data = new PowerPointChartData(
            new[] { "Q1", "Q2", "Q3", "Q4" },
            new[] {
                new PowerPointChartSeries("Sales", new[] { 12D, 18D, 24D, 30D }),
                new PowerPointChartSeries("Target", new[] { 15D, 20D, 22D, 28D })
            });
        PowerPointChart chart = presentation.AddSlide().AddChartPoints(data, 40, 32, 240, 172);
        chart.SetTitle("Revenue Mix");
        var options = new PowerPointPdfSaveOptions {
            ChartLayout = new OfficeChartLayout(preventLabelOverlap: false)
        };

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("40 36 240 172 re", raw, StringComparison.Ordinal);
        Assert.Contains("0.122 0.306 0.475 rg", raw, StringComparison.Ordinal);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Revenue Mix", text, StringComparison.Ordinal);
        Assert.Contains("Sales", text, StringComparison.Ordinal);
        Assert.Contains("Target", text, StringComparison.Ordinal);
        Assert.Contains("Q1", text, StringComparison.Ordinal);
        Assert.Contains("Q4", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersInheritedLayoutChartsFromLayoutPart() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 240);
        PowerPointSlide slide = presentation.AddSlide();
        var data = new PowerPointChartData(
            new[] { "Q1", "Q2", "Q3" },
            new[] {
                new PowerPointChartSeries("Sales", new[] { 12D, 18D, 24D })
            });
        PowerPointChart chart = slide.AddChartPoints(data, 40, 32, 240, 172);
        chart.SetTitle("Layout Revenue");
        ChartPart chartPart = GetChartPart(chart);
        SlideLayoutPart layoutPart = slide.SlidePart.SlideLayoutPart!;
        DocumentFormat.OpenXml.Presentation.GraphicFrame slideFrame = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<DocumentFormat.OpenXml.Presentation.GraphicFrame>()
            .Single(frame => frame.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>() != null);
        DocumentFormat.OpenXml.Presentation.GraphicFrame layoutFrame =
            (DocumentFormat.OpenXml.Presentation.GraphicFrame)slideFrame.CloneNode(true);
        layoutPart.AddPart(chartPart, "rIdLayoutChart");
        layoutFrame.Graphic!.GraphicData!.GetFirstChild<C.ChartReference>()!.Id = "rIdLayoutChart";
        layoutPart.SlideLayout.CommonSlideData!.ShapeTree!.Append(layoutFrame);
        chart.Remove();
        layoutPart.SlideLayout.Save();

        byte[] bytes = presentation.ToPdf();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Layout Revenue", text, StringComparison.Ordinal);
        Assert.Contains("Sales", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesHorizontalStackedBarChartKind() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 240);
        var data = new PowerPointChartData(
            new[] { "North", "South" },
            new[] {
                new PowerPointChartSeries("Won", new[] { 10D, 12D }),
                new PowerPointChartSeries("Open", new[] { 4D, 6D })
            });
        PowerPointChart chart = presentation.AddSlide().AddChartPoints(data, 40, 32, 240, 172);
        SetBarChartShape(chart, C.BarDirectionValues.Bar, C.BarGroupingValues.Stacked);

        Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot snapshot));
        Assert.Equal(PowerPointChartSnapshotKind.StackedBar, snapshot.ChartKind);

        byte[] bytes = presentation.ToPdf();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Won", text, StringComparison.Ordinal);
        Assert.Contains("Open", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesStackedLineChartKind() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 240);
        var data = new PowerPointChartData(
            new[] { "Q1", "Q2", "Q3" },
            new[] {
                new PowerPointChartSeries("Actual", new[] { 10D, 12D, 16D }),
                new PowerPointChartSeries("Target", new[] { 3D, 4D, 5D })
            });
        PowerPointChart chart = presentation.AddSlide().AddLineChartPoints(data, 40, 32, 240, 172);
        SetLineChartGrouping(chart, C.GroupingValues.PercentStacked);

        Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot snapshot));
        Assert.Equal(PowerPointChartSnapshotKind.StackedLine100, snapshot.ChartKind);

        byte[] bytes = presentation.ToPdf();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Actual", text, StringComparison.Ordinal);
        Assert.Contains("Target", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersAreaChartSnapshots() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 240);
        var data = new PowerPointChartData(
            new[] { "Q1", "Q2", "Q3" },
            new[] {
                new PowerPointChartSeries("Actual", new[] { 10D, 12D, 16D }),
                new PowerPointChartSeries("Target", new[] { 3D, 4D, 5D })
            });
        PowerPointChart chart = presentation.AddSlide().AddLineChartPoints(data, 40, 32, 240, 172);
        ConvertLineChartToAreaChart(chart, C.GroupingValues.Stacked);

        Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot snapshot));
        Assert.Equal(PowerPointChartSnapshotKind.StackedArea, snapshot.ChartKind);

        byte[] bytes = presentation.ToPdf();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Actual", text, StringComparison.Ordinal);
        Assert.Contains("Target", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersRadarChartSnapshots() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 240);
        var data = new PowerPointChartData(
            new[] { "Speed", "Quality", "Reach" },
            new[] {
                new PowerPointChartSeries("Actual", new[] { 10D, 12D, 16D }),
                new PowerPointChartSeries("Target", new[] { 8D, 11D, 14D })
            });
        PowerPointChart chart = presentation.AddSlide().AddLineChartPoints(data, 40, 32, 240, 172);
        ConvertLineChartToRadarChart(chart);

        Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot snapshot));
        Assert.Equal(PowerPointChartSnapshotKind.Radar, snapshot.ChartKind);

        byte[] bytes = presentation.ToPdf();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Actual", text, StringComparison.Ordinal);
        Assert.Contains("Target", text, StringComparison.Ordinal);
        Assert.Contains("Speed", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_PowerPointPresentation_RendersComboChartsInsteadOfWarning() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 240);
        var data = new PowerPointChartData(
            new[] { "Q1", "Q2" },
            new[] {
                new PowerPointChartSeries("Bars", new[] { 10D, 12D }),
                new PowerPointChartSeries("Trend", new[] { 11D, 13D })
            });
        PowerPointChart chart = presentation.AddSlide().AddChartPoints(data, 40, 32, 240, 172);
        ConvertSecondBarSeriesToLineChart(chart);
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.ToPdfDocument(options).ToBytes();

        Assert.Empty(options.Warnings);
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Bars", text, StringComparison.Ordinal);
        Assert.Contains("Trend", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersZeroThicknessLineAutoShapes() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddShapePoints(ShapeTypeValues.Line, 20, 40, 100, 0).Stroke("1E5A96", 1.5D);
        slide.AddShapePoints(ShapeTypeValues.Line, 140, 30, 0, 80).Stroke("C00000", 1.5D);
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("20 120 m", raw, StringComparison.Ordinal);
        Assert.Contains("120 120 l", raw, StringComparison.Ordinal);
        Assert.Contains("140 130 m", raw, StringComparison.Ordinal);
        Assert.Contains("140 50 l", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersZeroThicknessStraightConnectors() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddShapePoints(ShapeTypeValues.StraightConnector1, 20, 40, 100, 0).Stroke("1E5A96", 1.5D);
        slide.AddShapePoints(ShapeTypeValues.StraightConnector1, 140, 30, 0, 80).Stroke("C00000", 1.5D);
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("20 120 m", raw, StringComparison.Ordinal);
        Assert.Contains("120 120 l", raw, StringComparison.Ordinal);
        Assert.Contains("140 130 m", raw, StringComparison.Ordinal);
        Assert.Contains("140 50 l", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesFlippedLineAutoShapes() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointAutoShape line = presentation.AddSlide().AddShapePoints(ShapeTypeValues.Line, 20, 40, 80, 40);
        line.HorizontalFlip = true;
        line.Stroke("1E5A96", 1.5D);

        byte[] bytes = presentation.ToPdf();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("100 120 m", raw, StringComparison.Ordinal);
        Assert.Contains("20 80 l", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesScatterSeriesXValues() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 240);
        var data = new PowerPointScatterChartData(new[] {
            new PowerPointScatterChartSeries("Actual", new[] { 1D, 2D, 3D }, new[] { 10D, 12D, 14D }),
            new PowerPointScatterChartSeries("Forecast", new[] { 1.5D, 2.5D }, new[] { 11D, 13D })
        });
        PowerPointChart chart = presentation.AddSlide().AddScatterChartPoints(data, 40, 32, 240, 172);

        Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot snapshot));
        PowerPointChartSeries forecast = Assert.Single(snapshot.Data.Series, series => series.Name == "Forecast");
        Assert.Equal(new[] { 1.5D, 2.5D }, forecast.XValues);

        byte[] bytes = presentation.ToPdf();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Actual", text, StringComparison.Ordinal);
        Assert.Contains("Forecast", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_UsesFirstRenderableScatterSeriesForCategories() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 240);
        var data = new PowerPointScatterChartData(new[] {
            new PowerPointScatterChartSeries("EmptyX", new[] { 1D, 2D, 3D }, new[] { 10D, 12D, 14D }),
            new PowerPointScatterChartSeries("Forecast", new[] { 1.5D, 2.5D }, new[] { 11D, 13D })
        });
        PowerPointChart chart = presentation.AddSlide().AddScatterChartPoints(data, 40, 32, 240, 172);
        RemoveFirstScatterSeriesXValues(chart);

        Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot snapshot));

        Assert.Equal(new[] { "1.5", "2.5" }, snapshot.Data.Categories);
        PowerPointChartSeries forecast = Assert.Single(snapshot.Data.Series, series => series.Name == "Forecast");
        Assert.Equal(new[] { 1.5D, 2.5D }, forecast.XValues);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesSparseCachedChartPointIndices() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 240);
        var data = new PowerPointChartData(
            new[] { "Jan", "Feb", "Mar" },
            new[] {
                new PowerPointChartSeries("Actual", new[] { 10D, 20D, 30D })
            });
        PowerPointChart chart = presentation.AddSlide().AddChartPoints(data, 40, 32, 240, 172);
        MakeFirstBarSeriesCacheSparse(chart);

        Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot snapshot));

        Assert.Equal(new[] { "Jan", string.Empty, "Mar" }, snapshot.Data.Categories);
        PowerPointChartSeries actual = Assert.Single(snapshot.Data.Series);
        Assert.Equal(new[] { 10D, 0D, 30D }, actual.Values);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_ReadsCategoriesFromFirstRenderableCategorySeries() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 240);
        var data = new PowerPointChartData(
            new[] { "Jan", "Feb", "Mar" },
            new[] {
                new PowerPointChartSeries("Empty", new[] { 0D, 0D, 0D }),
                new PowerPointChartSeries("Actual", new[] { 10D, 20D, 30D })
            });
        PowerPointChart chart = presentation.AddSlide().AddChartPoints(data, 40, 32, 240, 172);
        MakeFirstBarSeriesCacheEmpty(chart);

        Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot snapshot));

        Assert.Equal(new[] { "Jan", "Feb", "Mar" }, snapshot.Data.Categories);
        PowerPointChartSeries actual = Assert.Single(snapshot.Data.Series, series => series.Name == "Actual");
        Assert.Equal(new[] { 10D, 20D, 30D }, actual.Values);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_AppliesSharedChartStyleAndLayoutOptions() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(360, 240);
        var data = new PowerPointChartData(
            new[] { "M01", "M02", "M03", "M04", "M05", "M06", "M07", "M08" },
            new[] {
                new PowerPointChartSeries("Actual", new[] { 12D, 18D, 24D, 30D, 34D, 38D, 41D, 43D }),
                new PowerPointChartSeries("Target", new[] { 15D, 20D, 22D, 28D, 31D, 36D, 39D, 44D })
            });
        PowerPointChart chart = presentation.AddSlide().AddChartPoints(data, 38, 30, 270, 176);
        chart.SetTitle("Styled Slide Chart");
        var options = new PowerPointPdfSaveOptions {
            ChartStyle = new OfficeChartStyle(
                palette: new[] {
                    OfficeColor.FromRgb(18, 52, 86),
                    OfficeColor.FromRgb(120, 40, 160)
                },
                backgroundColor: OfficeColor.FromRgb(242, 248, 255),
                titleColor: OfficeColor.FromRgb(200, 10, 10)),
            ChartLayout = new OfficeChartLayout(maximumCategoryAxisLabels: 2)
        };

        byte[] bytes = presentation.ToPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.071 0.204 0.337 rg", raw, StringComparison.Ordinal);
        Assert.Contains("0.471 0.157 0.627 rg", raw, StringComparison.Ordinal);
        Assert.Contains("0.949 0.973 1 rg", raw, StringComparison.Ordinal);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Styled Slide Chart", text, StringComparison.Ordinal);
        Assert.Contains("Actual", text, StringComparison.Ordinal);
        Assert.Contains("Target", text, StringComparison.Ordinal);
        Assert.Contains("M01", text, StringComparison.Ordinal);
        Assert.Contains("M05", text, StringComparison.Ordinal);
        Assert.DoesNotContain("M02", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_WarnsWhenSharedChartQualityPreflightFindsIssues() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 220);
        var data = new PowerPointChartData(
            new[] { "M01", "M02", "M03", "M04", "M05", "M06", "M07", "M08", "M09", "M10", "M11", "M12" },
            new[] {
                new PowerPointChartSeries("Actual", new[] { 12D, 18D, 24D, 30D, 34D, 38D, 41D, 43D, 44D, 45D, 46D, 47D })
            });
        PowerPointChart chart = presentation.AddSlide().AddChartPoints(data, 32, 28, 240, 150);
        chart.SetTitle("Dense Slide Chart");
        var options = new PowerPointPdfSaveOptions {
            ChartLayout = new OfficeChartLayout(maximumCategoryAxisLabels: 12, preventLabelOverlap: false)
        };

        PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult(options);
        byte[] bytes = result.ToBytes();

        PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings, item => item.Code == "chart-quality");
        Assert.Equal("Slide 1", warning.Source);
        Assert.Contains("Dense Slide Chart", warning.Message, StringComparison.Ordinal);
        Assert.Contains("TextOverlap", warning.Message, StringComparison.Ordinal);
        Assert.NotNull(warning.LayoutDiagnostic);
        Assert.Equal(PdfCore.PdfLayoutDiagnosticKind.SimplifiedContent, warning.LayoutDiagnostic!.Kind);
        Assert.Equal("PowerPointChart", warning.LayoutDiagnostic.Source);
        Assert.True(warning.LayoutDiagnostic.HasBounds);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Contains("Dense Slide Chart", pdf.GetPage(1).Text, StringComparison.Ordinal);
    }

    private static byte[] CreateMinimalRgbPng() => PdfPngTestImages.CreateRgbPng(255, 0, 0);

    private static void SetBarChartShape(PowerPointChart chart, C.BarDirectionValues direction, C.BarGroupingValues grouping) {
        ChartPart chartPart = GetChartPart(chart);
        C.BarChart barChart = chartPart.ChartSpace!.Descendants<C.BarChart>().Single();
        barChart.GetFirstChild<C.BarDirection>()!.Val = direction;
        barChart.GetFirstChild<C.BarGrouping>()!.Val = grouping;
        chartPart.ChartSpace.Save();
    }

    private static void SetLineChartGrouping(PowerPointChart chart, C.GroupingValues grouping) {
        ChartPart chartPart = GetChartPart(chart);
        C.LineChart lineChart = chartPart.ChartSpace!.Descendants<C.LineChart>().Single();
        C.Grouping chartGrouping = lineChart.GetFirstChild<C.Grouping>() ?? lineChart.PrependChild(new C.Grouping());
        chartGrouping.Val = grouping;
        chartPart.ChartSpace.Save();
    }

    private static void ConvertLineChartToAreaChart(PowerPointChart chart, C.GroupingValues grouping) {
        ChartPart chartPart = GetChartPart(chart);
        C.PlotArea plotArea = chartPart.ChartSpace!.Descendants<C.PlotArea>().Single();
        C.LineChart lineChart = plotArea.Elements<C.LineChart>().Single();
        var areaChart = new C.AreaChart(new C.Grouping { Val = grouping });
        foreach (C.LineChartSeries lineSeries in lineChart.Elements<C.LineChartSeries>()) {
            var areaSeries = new C.AreaChartSeries();
            foreach (OpenXmlElement child in lineSeries.ChildElements) {
                areaSeries.Append(child.CloneNode(true));
            }

            areaChart.Append(areaSeries);
        }

        foreach (C.AxisId axisId in lineChart.Elements<C.AxisId>()) {
            areaChart.Append((C.AxisId)axisId.CloneNode(true));
        }

        lineChart.InsertAfterSelf(areaChart);
        lineChart.Remove();
        chartPart.ChartSpace.Save();
    }

    private static void ConvertSecondBarSeriesToLineChart(PowerPointChart chart) {
        ChartPart chartPart = GetChartPart(chart);
        C.PlotArea plotArea = chartPart.ChartSpace!.Descendants<C.PlotArea>().Single();
        C.BarChart barChart = plotArea.Elements<C.BarChart>().Single();
        C.BarChartSeries barSeries = barChart.Elements<C.BarChartSeries>().Skip(1).Single();
        var lineChart = new C.LineChart(new C.Grouping { Val = C.GroupingValues.Standard });
        var lineSeries = new C.LineChartSeries();
        foreach (OpenXmlElement child in barSeries.ChildElements) {
            lineSeries.Append(child.CloneNode(true));
        }

        lineChart.Append(lineSeries);
        foreach (C.AxisId axisId in barChart.Elements<C.AxisId>()) {
            lineChart.Append((C.AxisId)axisId.CloneNode(true));
        }

        barSeries.Remove();
        barChart.InsertAfterSelf(lineChart);
        chartPart.ChartSpace.Save();
    }

    private static void ConvertLineChartToRadarChart(PowerPointChart chart) {
        ChartPart chartPart = GetChartPart(chart);
        C.PlotArea plotArea = chartPart.ChartSpace!.Descendants<C.PlotArea>().Single();
        C.LineChart lineChart = plotArea.Elements<C.LineChart>().Single();
        var radarChart = new C.RadarChart(new C.RadarStyle { Val = C.RadarStyleValues.Standard });
        foreach (C.LineChartSeries lineSeries in lineChart.Elements<C.LineChartSeries>()) {
            var radarSeries = new C.RadarChartSeries();
            foreach (OpenXmlElement child in lineSeries.ChildElements) {
                radarSeries.Append(child.CloneNode(true));
            }

            radarChart.Append(radarSeries);
        }

        foreach (C.AxisId axisId in lineChart.Elements<C.AxisId>()) {
            radarChart.Append((C.AxisId)axisId.CloneNode(true));
        }

        lineChart.InsertAfterSelf(radarChart);
        lineChart.Remove();
        chartPart.ChartSpace.Save();
    }

    private static void MakeFirstBarSeriesCacheSparse(PowerPointChart chart) {
        ChartPart chartPart = GetChartPart(chart);
        C.BarChartSeries series = chartPart.ChartSpace!.Descendants<C.BarChartSeries>().Single();
        C.StringCache categoryCache = series.GetFirstChild<C.CategoryAxisData>()!.Descendants<C.StringCache>().Single();
        C.NumberingCache valueCache = series.GetFirstChild<C.Values>()!.Descendants<C.NumberingCache>().Single();
        categoryCache.Elements<C.StringPoint>().Single(point => point.Index?.Value == 1U).Remove();
        valueCache.Elements<C.NumericPoint>().Single(point => point.Index?.Value == 1U).Remove();
        chartPart.ChartSpace.Save();
    }

    private static void MakeFirstBarSeriesCacheEmpty(PowerPointChart chart) {
        ChartPart chartPart = GetChartPart(chart);
        C.BarChartSeries series = chartPart.ChartSpace!.Descendants<C.BarChartSeries>().First();
        foreach (C.StringPoint point in series.GetFirstChild<C.CategoryAxisData>()!.Descendants<C.StringPoint>().ToList()) {
            point.Remove();
        }

        foreach (C.NumericPoint point in series.GetFirstChild<C.Values>()!.Descendants<C.NumericPoint>().ToList()) {
            point.Remove();
        }

        chartPart.ChartSpace.Save();
    }

    private static void RemoveFirstScatterSeriesXValues(PowerPointChart chart) {
        ChartPart chartPart = GetChartPart(chart);
        C.ScatterChartSeries series = chartPart.ChartSpace!.Descendants<C.ScatterChartSeries>().First();
        series.GetFirstChild<C.XValues>()?.Remove();
        chartPart.ChartSpace.Save();
    }

    private static ChartPart GetChartPart(PowerPointChart chart) {
        MethodInfo method = typeof(PowerPointChart).GetMethod("GetChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (ChartPart)method.Invoke(chart, Array.Empty<object>())!;
    }

    private static double FindWordStartY(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].StartBaseLine.Y;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static double FindWordStartX(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].StartBaseLine.X;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static double AverageLetterFontSize(UglyToad.PdfPig.Content.Page page, string word) {
        var letters = page.Letters.ToList();
        string text = string.Join("", letters.Select(letter => letter.Value));
        int index = text.IndexOf(word, StringComparison.Ordinal);
        Assert.True(index >= 0, "Expected to find word '" + word + "' in PDF text '" + text + "'.");
        return letters.Skip(index).Take(word.Length).Average(letter => letter.FontSize);
    }

    private static void AssertRawPdfContainsAnyBaseFont(string rawPdf, params string[] fontNameParts) {
        Assert.True(
            fontNameParts.Any(fontNamePart => rawPdf.Contains("/BaseFont /" + fontNamePart, StringComparison.OrdinalIgnoreCase)),
            "Expected raw PDF to contain one of these BaseFont names: " + string.Join(", ", fontNameParts));
    }
}
