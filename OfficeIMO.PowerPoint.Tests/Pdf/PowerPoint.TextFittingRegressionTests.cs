using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointTextFittingRegressionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StoredNormalFitScalesAllRunsAndReducesOnlyPercentageLeading(bool absoluteLeading) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        var box = CreateBox(presentation, "FIRST\nSECOND", 220, 140);
        box.Paragraphs[0].Runs[0].FontSize = global::OfficeIMO.OfficeLength.Points(24);
        box.Paragraphs[0].Runs[0].SetHyperlink("https://officeimo.com/text-fitting");
        box.Paragraphs[0].AddRun("TAIL", run => { run.FontSize = global::OfficeIMO.OfficeLength.Points(16); run.Bold = true; });
        if (absoluteLeading) box.Paragraphs[0].LineSpacing = global::OfficeIMO.OfficeLength.Points(30);
        else box.Paragraphs[0].LineSpacingMultiplier = 1.2D;
        box.SetTextAutoFit(PowerPointTextAutoFit.Normal, new PowerPointTextAutoFitOptions(50, 20));

        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var drawing = Draw(box, 220, 140, diagnostics);
        var text = Assert.IsType<OfficeDrawingRichText>(Assert.Single(drawing.Elements));
        Assert.Equal(new[] { 12D, 8D }, text.Runs.Select(run => run.FontSize));
        var layout = OfficeDrawingTextLayout.Create(text, text.Width, text.Height,
            (value, size, _, _) => (value?.Length ?? 0) * size / 2D);
        Assert.All(layout.Lines, line => Assert.Equal(absoluteLeading ? 30D : 12D, line.LineHeight, 6));
        Assert.Equal("https://officeimo.com/text-fitting", text.Runs[0].LinkUri);
        Assert.True(text.Runs[1].Bold);
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == "powerpoint-text-overflow");
        Assert.Equal(24, box.Paragraphs[0].Runs[0].FontSize?.ToPoints());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MissingNormalFontScaleFitsCompleteParagraphFlow(bool storedReduction) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        var box = CreateBox(presentation, "FIRST", 120, 45);
        box.Paragraphs[0].SpaceAfter = global::OfficeIMO.OfficeLength.Points(3);
        box.AddParagraph("SECOND");
        box.AddParagraph("THIRD");
        box.SetTextAutoFit(PowerPointTextAutoFit.Normal,
            new PowerPointTextAutoFitOptions(lineSpaceReductionPercent: storedReduction ? 20 : null));
        string source = box.TextBody!.OuterXml;

        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var drawing = Draw(box, 120, 45, diagnostics);
        Assert.Equal(45D, drawing.Height);
        Assert.Equal(3, drawing.Elements.Count);
        Assert.All(drawing.Elements, element => {
            var text = Assert.IsType<OfficeDrawingText>(element);
            Assert.InRange(text.Font.Size, 1D, 23.99D);
            Assert.True(text.Y + text.Height <= 45.02D);
        });
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == "powerpoint-text-overflow");
        Assert.Equal(source, box.TextBody.OuterXml);
    }

    [Fact]
    public void MissingNormalScaleRewrapsToFitHeightWithoutShrinkingTextThatAlreadyFits() {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        var box = CreateBox(presentation, "ALPHA BETA GAMMA DELTA EPSILON ZETA", 110, 45);
        box.TextAutoFit = PowerPointTextAutoFit.Normal;
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var drawing = Draw(box, 110, 45, diagnostics);
        Assert.True(Assert.IsType<OfficeDrawingText>(Assert.Single(drawing.Elements)).Font.Size < 24D);
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == "powerpoint-text-overflow");

        var fits = CreateBox(presentation, "FITS", 220, 80);
        fits.TextAutoFit = PowerPointTextAutoFit.Normal;
        drawing = Draw(fits, 220, 80, new List<OfficeImageExportDiagnostic>());
        Assert.Equal(24D, Assert.IsType<OfficeDrawingText>(Assert.Single(drawing.Elements)).Font.Size);
    }

    [Theory]
    [InlineData(PowerPointTextAutoFit.None)]
    [InlineData(PowerPointTextAutoFit.Normal)]
    public void FixedFrameOverflowRetainsEveryParagraphAndReportsOverflowWithoutClaimingOmission(PowerPointTextAutoFit mode) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(400),OfficeIMO.OfficeLength.Points(300));
        var box = CreateBox(presentation, "FIRST", 220, 15);
        box.AddParagraph("SECOND");
        box.AddParagraph("LAST");
        box.FillColor = "1264A3";
        box.SetTextAutoFit(mode, mode == PowerPointTextAutoFit.Normal ? new PowerPointTextAutoFitOptions(100, 0) : null);

        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var drawing = Draw(box, 220, 15, diagnostics);
        Assert.Equal(15D, drawing.Height);
        Assert.Equal(new[] { "FIRST", "SECOND", "LAST" }, drawing.Elements.Cast<OfficeDrawingText>().Select(text => text.Text));
        Assert.All(drawing.Elements.Cast<OfficeDrawingText>(), text => Assert.Equal(24D, text.Font.Size));
        var overflow = Assert.Single(diagnostics, diagnostic => diagnostic.Code == "powerpoint-text-overflow");
        Assert.Equal(OfficeConversionLossKind.None, overflow.LossKind);
        Assert.True(drawing.Elements.Cast<OfficeDrawingText>().Last().Y > drawing.Height);

        var options = new PowerPointPdfExportOptions();
        var conversion = presentation.ConvertToPdfWithEvidence(options);
        byte[] bytes = conversion.ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Contains("FIRSTSECONDLAST", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("...", pdf.GetPage(1).Text, StringComparison.Ordinal);
        var warning = Assert.Single(conversion.Warnings, item => item.Code == "text-box-overflow");
        Assert.Equal(OfficeIMO.Pdf.PdfLayoutDiagnosticKind.Overflow, warning.LayoutDiagnostic!.Kind);
        Assert.Equal(OfficeConversionLossKind.None, warning.LossKind);
        var frame = Assert.Single(OfficeIMO.Pdf.PdfDocument.Load(new System.IO.MemoryStream(bytes)).Pages[0].ToDrawing().Shapes,
            shape => shape.Shape.FillColor == OfficeColor.FromRgb(18, 100, 163));
        Assert.InRange(frame.Shape.Width, 219.99D, 220.01D);
        Assert.InRange(frame.Shape.Height, 14.99D, 15.01D);
    }

    [Fact]
    public void ShapeFitGrowsExportedFrameAndKeepsSourceGeometryAndFontSizes() {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        var box = CreateBox(presentation, "FIRST", 220, 15);
        box.AddParagraph("SECOND");
        box.AddParagraph("LAST");
        box.FillColor = "1264A3";
        box.TextAutoFit = PowerPointTextAutoFit.Shape;
        string source = box.Element.OuterXml;
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var drawing = Draw(box, 220, 15, diagnostics);
        Assert.True(drawing.Height >= 87D);
        Assert.Equal(220D, drawing.Width);
        Assert.All(drawing.Elements.Cast<OfficeDrawingText>(), text => Assert.Equal(24D, text.Font.Size));
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == "powerpoint-text-overflow");
        Assert.Equal(source, box.Element.OuterXml);

        byte[] bytes = presentation.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Contains("FIRSTSECONDLAST", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.All(pdf.GetPage(1).Letters, letter => Assert.InRange(letter.FontSize, 23.9D, 24.1D));
        var frame = Assert.Single(OfficeIMO.Pdf.PdfDocument.Load(new System.IO.MemoryStream(bytes)).Pages[0].ToDrawing().Shapes,
            shape => shape.Shape.FillColor == OfficeColor.FromRgb(18, 100, 163));
        Assert.InRange(frame.Shape.Width, 219.99D, 220.01D);
        Assert.InRange(frame.Shape.Height, drawing.Height - 0.01D, drawing.Height + 0.01D);
    }

    [Fact]
    public void ShapeFitGrowsTheNativePdfEllipseFrameAfterTextMeasurement() {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        var box = presentation.AddSlide().AddTextShape(OfficePresetShapeType.Ellipse, "FIRST", OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(220), OfficeIMO.OfficeLength.Points(15));
        box.AddParagraph("SECOND");
        box.AddParagraph("LAST");
        box.FontName = "Helvetica";
        box.FontSize = global::OfficeIMO.OfficeLength.Points(24);
        box.FillColor = "1264A3";
        box.TextMarginLeft = box.TextMarginRight = box.TextMarginTop = box.TextMarginBottom = global::OfficeIMO.OfficeLength.Points(0);
        box.TextAutoFit = PowerPointTextAutoFit.Shape;
        string source = box.Element.OuterXml;

        byte[] bytes = presentation.ConvertToPdfWithEvidence().ToBytes();
        var page = OfficeIMO.Pdf.PdfDocument.Load(new System.IO.MemoryStream(bytes)).Pages[0];
        var frame = Assert.Single(page.ToDrawing().Shapes,
            shape => shape.Shape.FillColor == OfficeColor.FromRgb(18, 100, 163));
        Assert.InRange(frame.Shape.Width, 219.99D, 220.01D);
        Assert.InRange(frame.Shape.Height, 86.99D, 87.01D);
        Assert.Equal("FIRSTSECONDLAST", string.Concat(page.GetTextSpans().Select(span => span.Text)));
        Assert.Equal(source, box.Element.OuterXml);

        var geometryOnly = presentation.ConvertToPdfWithEvidence(new PowerPointPdfExportOptions { IncludeTextBoxes = false });
        var geometryPage = OfficeIMO.Pdf.PdfReadDocument.Open(geometryOnly.ToBytes()).Pages[0];
        var authoredFrame = Assert.Single(geometryPage.ToDrawing().Shapes,
            shape => shape.Shape.FillColor == OfficeColor.FromRgb(18, 100, 163));
        Assert.InRange(authoredFrame.Shape.Height, 14.99D, 15.01D);
        Assert.Empty(geometryPage.GetTextSpans());
    }

    [Theory]
    [InlineData(PowerPointTextAutoFit.None)]
    [InlineData(PowerPointTextAutoFit.Normal)]
    [InlineData(PowerPointTextAutoFit.Shape)]
    public void UnwrappedTextStaysOnOneLineAndShapeFitGrowsWidth(PowerPointTextAutoFit mode) {
        const string value = "UNWRAPPED TEXT STAYS COMPLETE";
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(600),OfficeIMO.OfficeLength.Points(200));
        var box = CreateBox(presentation, value, 80, 60);
        box.FillColor = "1264A3";
        box.TextAutoFit = mode;
        box.TextBody!.GetFirstChild<A.BodyProperties>()!.Wrap = A.TextWrappingValues.None;
        string source = box.Element.OuterXml;

        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var drawing = Draw(box, 80, 60, diagnostics);
        var text = Assert.IsType<OfficeDrawingRichText>(Assert.Single(drawing.Elements));
        Assert.False(text.WrapText);
        Assert.Equal(value, text.PlainText);
        Assert.Equal(60D, drawing.Height);
        Assert.Equal(mode == PowerPointTextAutoFit.None,
            diagnostics.Any(diagnostic => diagnostic.Code == "powerpoint-text-overflow"));
        if (mode == PowerPointTextAutoFit.Shape) Assert.True(drawing.Width > 80D);
        else Assert.Equal(80D, drawing.Width);
        if (mode == PowerPointTextAutoFit.Normal) Assert.InRange(text.Runs[0].FontSize, 1D, 23.99D);
        else Assert.Equal(24D, text.Runs[0].FontSize);

        byte[] bytes = presentation.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        var page = OfficeIMO.Pdf.PdfDocument.Load(new System.IO.MemoryStream(bytes)).Pages[0];
        var spans = page.GetTextSpans();
        Assert.Equal(value, string.Concat(spans.Select(span => span.Text)));
        Assert.All(spans, span => {
            Assert.True(span.IsVisible);
            Assert.InRange(Math.Abs(span.Y - spans[0].Y), 0D, 0.01D);
        });
        var frame = Assert.Single(page.ToDrawing().Shapes,
            shape => shape.Shape.FillColor == OfficeColor.FromRgb(18, 100, 163));
        Assert.InRange(frame.Shape.Height, 59.99D, 60.01D);
        if (mode == PowerPointTextAutoFit.Shape) Assert.True(frame.Shape.Width > 80D);
        else Assert.InRange(frame.Shape.Width, 79.99D, 80.01D);
        Assert.Equal(source, box.Element.OuterXml);
    }

    [Fact]
    public void UnwrappedShapeFitGrowsBothDimensionsForAuthoredHardBreaks() {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(600),OfficeIMO.OfficeLength.Points(200));
        var box = CreateBox(presentation, "FIRST LONG LINE\nSECOND LONG LINE", 80, 10);
        box.FillColor = "1264A3";
        box.TextAutoFit = PowerPointTextAutoFit.Shape;
        box.TextBody!.GetFirstChild<A.BodyProperties>()!.Wrap = A.TextWrappingValues.None;
        string source = box.Element.OuterXml;
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var drawing = Draw(box, 80, 10, diagnostics);
        Assert.True(drawing.Width > 80D);
        Assert.InRange(drawing.Height, 57.59D, 57.61D);
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == "powerpoint-text-overflow");

        var page = OfficeIMO.Pdf.PdfReadDocument.Open(presentation.ConvertToPdfWithEvidence().ToBytes()).Pages[0];
        var spans = page.GetTextSpans();
        var first = Assert.Single(spans, span => span.Text.Contains("FIRST LONG LINE", StringComparison.Ordinal));
        var second = Assert.Single(spans, span => span.Text.Contains("SECOND LONG LINE", StringComparison.Ordinal));
        Assert.InRange(first.Y - second.Y, 28.79D, 28.81D);
        Assert.All(spans, span => Assert.Equal(24D, span.FontSize));
        var frame = Assert.Single(page.ToDrawing().Shapes,
            shape => shape.Shape.FillColor == OfficeColor.FromRgb(18, 100, 163));
        Assert.True(frame.Shape.Width > 80D);
        Assert.InRange(frame.Shape.Height, 57.59D, 57.61D);
        Assert.Equal(source, box.Element.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ComputedFittingObservesCancellationDuringTextMeasurement(bool unwrapped) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        var box = CreateBox(presentation, "FIRST SECOND THIRD FOURTH FIFTH LAST", 40, 10);
        box.TextAutoFit = PowerPointTextAutoFit.Normal;
        if (unwrapped) box.TextBody!.GetFirstChild<A.BodyProperties>()!.Wrap = A.TextWrappingValues.None;
        using var cancellation = new CancellationTokenSource();
        int measurements = 0;
        Assert.Throws<OperationCanceledException>(() => PowerPointSlideImageRenderer.CreateTextBoxDrawing(
            box, 40, 10, new List<OfficeImageExportDiagnostic>(), (text, size, _, _) => {
                measurements++;
                cancellation.Cancel();
                return (text?.Length ?? 0) * size * 0.5D;
            }, "Helvetica", cancellation.Token));
        Assert.True(measurements > 0);
    }

    [Fact]
    public void PartlyOffLeftTextBoxRasterRetainsOverflowBeyondItsOriginalFrame() {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(160),OfficeIMO.OfficeLength.Points(100));
        var slide = presentation.AddSlide();
        var box = slide.AddTextBox("OOOOOOOO",global::OfficeIMO.OfficeLength.Points(-12),global::OfficeIMO.OfficeLength.Points(10),global::OfficeIMO.OfficeLength.Points(24),global::OfficeIMO.OfficeLength.Points(10));
        box.AddParagraph("OOOOOOOO");
        box.FontName = "OverflowRaster";
        box.FontSize = global::OfficeIMO.OfficeLength.Points(24);
        box.Color = "C02020";
        box.TextMarginLeft = box.TextMarginRight = box.TextMarginTop = box.TextMarginBottom = global::OfficeIMO.OfficeLength.Points(0);
        box.TextAutoFit = PowerPointTextAutoFit.None;
        box.TextBody!.GetFirstChild<A.BodyProperties>()!.Wrap = A.TextWrappingValues.None;
        var fonts = new OfficeFontFaceCollection().Add("OverflowRaster",
            OfficeIMO.TestAssets.PdfTestFontAssets.LoadBundledOpenTypeCffFont());
        var options = new PowerPointImageExportOptions {
            Scale = 1D, MaximumRasterPixels = 16000, IncludeSlideBackground = false,
            BackgroundColor = OfficeColor.White
        };
        options.UseRenderingProfile(new OfficeRenderingProfile("off-left-text-overflow", fonts));

        var result = slide.ExportImage(OfficeImageExportFormat.Png, options);
        Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
        Assert.Equal(160, image!.Width);
        Assert.Equal(100, image.Height);
        int paintedOverflow = 0;
        // This region is below and to the right of the original frame's visible
        // intersection with the slide (x=0..12, y=10..20).
        for (int y = 40; y < 80; y++) {
            for (int x = 24; x < 140; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                if (pixel.R > 150 && pixel.G < 80 && pixel.B < 80 && pixel.A > 0) paintedOverflow++;
            }
        }
        Assert.True(paintedOverflow > 10, "The slide clip discarded text outside the original text box frame.");
    }

    private static PowerPointTextBox CreateBox(PowerPointDocument presentation, string text, double width, double height) {
        var box = presentation.AddSlide().AddTextBox(text,global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(width),global::OfficeIMO.OfficeLength.Points(height));
        box.FontName = "Helvetica";
        box.FontSize = global::OfficeIMO.OfficeLength.Points(24);
        box.TextMarginLeft = box.TextMarginRight = box.TextMarginTop = box.TextMarginBottom = global::OfficeIMO.OfficeLength.Points(0);
        return box;
    }

    private static OfficeDrawing Draw(PowerPointTextBox box, double width, double height, List<OfficeImageExportDiagnostic> diagnostics) =>
        PowerPointSlideImageRenderer.CreateTextBoxDrawing(box, width, height, diagnostics,
            (text, size, _, _) => (text?.Length ?? 0) * size * 0.5D, "Helvetica");
}
