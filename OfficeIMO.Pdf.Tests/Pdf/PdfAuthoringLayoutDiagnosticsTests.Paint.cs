using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfAuthoringLayoutDiagnosticsTests {
    [Theory]
    [InlineData("flow", false)]
    [InlineData("column", false)]
    [InlineData("canvas", false)]
    [InlineData("flow", true)]
    [InlineData("column", true)]
    [InlineData("canvas", true)]
    public void InlineBorderOverhangUsesPaintInsteadOfItsLayoutAdvance(string frame, bool border) {
        var report = new PdfConversionReport();
        var box = new PdfInlineBox(120, 8, background: PdfColor.Black,
            borderColor: border ? PdfColor.Black : null, borderWidth: 2);
        var cell = new PdfTableCell(new[] { PdfTextRun.Inline(box) });
        PdfTableStyle style = CreateTableStyle(18);
        style.ClipTextToCellBounds = true;
        byte[] bytes = RenderPaintTable(frame, CreateOptions(report), cell, style);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Equal(border, report.Warnings.Any(warning => warning.Code == "table-cell-paint-clipped"));
        Assert.Equal(border, report.HasLoss);
    }

    [Theory]
    [InlineData("flow", -90, false)]
    [InlineData("column", -90, true)]
    [InlineData("canvas", -90, true)]
    [InlineData("flow", 90, true)]
    [InlineData("column", 90, false)]
    [InlineData("canvas", 90, false)]
    public void QuarterTurnDiagnosticsUseThePhysicalCellClip(string frame, int rotation, bool overflow) {
        var report = new PdfConversionReport();
        var cell = new PdfTableCell(new[] { PdfTextRun.Normal(overflow ? new string('W', 100) : "A") },
            paragraphs: null, textRotation: rotation);
        PdfTableStyle style = CreateTableStyle(18);
        style.ClipTextToCellBounds = true;
        using PdfPigDocument pdf = PdfPigDocument.Open(RenderPaintTable(frame, CreateOptions(report), cell, style));
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Equal(overflow, report.Warnings.Any(warning => warning.Code == "table-cell-paint-clipped"));
        Assert.Equal(overflow, report.HasLoss);
    }

    [Theory]
    [InlineData("flow", false)]
    [InlineData("column", false)]
    [InlineData("canvas", false)]
    [InlineData("flow", true)]
    [InlineData("column", true)]
    [InlineData("canvas", true)]
    public void ViewportChecksPaintBeforeInvisibleLineSlotsAreOmitted(string frame, bool hidden) {
        var report = new PdfConversionReport();
        var cell = new PdfTableCell("A").WithViewport(new PdfTableCellViewport(hidden ? 240 : 120, 18, 120, 18,
            offsetX: hidden ? 120 : 0));
        using PdfPigDocument pdf = PdfPigDocument.Open(RenderPaintTable(frame, CreateOptions(report), cell, CreateTableStyle(18)));
        Assert.Equal(hidden, report.Warnings.Any(warning => warning.Code == "table-cell-paint-clipped"));
        Assert.Equal(hidden, report.HasLoss);
        if (!hidden) Assert.Contains("A", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void CellParagraphFramesHonorTheEffectiveExpandedClip(bool canvas, bool strict) {
        var report = new PdfConversionReport();
        var runs = new[] { PdfTextRun.Normal("AAAA") };
        var cell = new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs, leftIndent: -6) });
        PdfTableStyle style = CreateTableStyle(18);
        style.ClipTextToCellBounds = strict;
        using PdfPigDocument pdf = PdfPigDocument.Open(RenderPaintTable(canvas ? "canvas" : "flow", CreateOptions(report), cell, style));
        // Flow paragraphs deliberately expand non-strict clips. Fixed canvas cells retain their own clip.
        Assert.Equal(canvas || strict, report.Warnings.Any(warning => warning.Code == "table-cell-paint-clipped"));
        Assert.Equal(canvas || strict, report.HasLoss);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StandardGlyphOverhangIsIndependentOfAdvance(bool tight) {
        var report = new PdfConversionReport();
        var runs = new[] { PdfTextRun.Normal("T", fontSize: 9, font: PdfStandardFont.TimesItalic) };
        using PdfPigDocument pdf = PdfPigDocument.Open(PdfDocument.Create(CreateOptions(report)).Canvas(canvas =>
            canvas.Text(runs, 24, 24, tight ? 5.004D : 8D, 18)).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Contains("T", pdf.GetPage(1).Text);
        Assert.Equal(tight, report.Warnings.Any(warning => warning.Code == "canvas-text-paint-clipped"));
        Assert.Equal(tight, report.HasLoss);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FreeTextReportsOmittedPaintWithoutTreatingUnusedLeadingAsLoss(bool overflow) {
        var report = new PdfConversionReport();
        byte[] bytes = PdfDocument.Create(CreateOptions(report)).Canvas(canvas =>
            canvas.FreeTextAnnotation(overflow ? "First\nSecond" : "First", 24, 24, 120, 18,
                fontSize: 9, padding: 0, lineHeight: overflow ? 12 : 108)).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.Single(PdfInspector.Inspect(bytes).GetAnnotationsBySubtype("FreeText"));
        Assert.Equal(overflow, report.Warnings.Any(warning => warning.Code == "free-text-annotation-content-omitted"));
        Assert.Equal(overflow, report.HasLoss);
    }

    [Theory]
    [InlineData(0D, 1D, false)]
    [InlineData(1D, 0D, false)]
    [InlineData(1D, 1D, true)]
    public void ZeroAreaOrTransparentEffectsHaveNoClippedPaint(double scale, double opacity, bool expected) {
        var report = new PdfConversionReport();
        byte[] bytes = PdfDocument.Create(CreateOptions(report)).Canvas(canvas => canvas.Effect(
            new OfficeTransform(scale, 0, 0, scale, 0, 0), opacity, effect => effect.TextBox(
                new[] { PdfTextRun.Inline(new PdfInlineBox(120, 8, borderColor: PdfColor.Black, borderWidth: 2)) },
                24, 24, 120, 18, new PdfCanvasTextBoxStyle { FontSize = 9, PaddingX = 0, PaddingY = 0 }))).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.NotEmpty(bytes);
        Assert.Equal(expected, report.Warnings.Any(warning => warning.Code == "canvas-text-box-paint-clipped"));
    }

    [Fact]
    public void TransformCanMoveLocalOffPageInkOntoThePhysicalPage() {
        var report = new PdfConversionReport();
        byte[] bytes = PdfDocument.Create(CreateOptions(report)).Canvas(canvas =>
            canvas.Effect(OfficeTransform.Translate(100, 0), 1D, effect => effect.Text("Fits", -80, 24, 60, 18))).ToBytes();
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("Fits", pdf.GetPage(1).Text);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "canvas-text-page-bounds-paint-clipped");
        Assert.False(report.HasLoss);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    [InlineData("canvas")]
    [InlineData("textbox")]
    [InlineData("annotation")]
    public void EmbeddedFontLeadersAndOverhangProduceIdenticalPdfBytesWithDiagnostics(string frame) {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Typography", "OfficeIMOBaselineSans-Regular.ttf"));
        var report = new PdfConversionReport();
        byte[] Render(bool diagnose) {
            var options = new PdfCreateOptions {
                PageWidth = 260, PageHeight = 180, MarginLeft = 24, MarginRight = 24, MarginTop = 24, MarginBottom = 24,
                DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 9,
                CompressEmbeddedFonts = false, CompressContentStreams = false
            }.EmbedStandardFont(PdfStandardFont.Helvetica, font, "Diagnostic subset probe");
            if (diagnose) options.ReportDiagnosticsTo(report, "SubsetProbe");
            var runs = new[] { PdfTextRun.Normal("ABC"), PdfTextRun.Tab(PdfTabLeaderStyle.Dots),
                PdfTextRun.Normal("Z", backgroundColor: PdfColor.Gray),
                PdfTextRun.Inline(new PdfInlineBox(120, 8, borderColor: PdfColor.Black, borderWidth: 2)) };
            if (frame == "annotation") return PdfDocument.Create(options).Canvas(canvas =>
                canvas.FreeTextAnnotation("ABC\nZ", 24, 24, 120, 10, fontSize: 9, padding: 0)).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
            if (frame == "textbox") return PdfDocument.Create(options).Canvas(canvas => canvas.TextBox(runs, 24, 24, 120, 18,
                new PdfCanvasTextBoxStyle { FontSize = 9, PaddingX = 0, PaddingY = 0 })).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
            PdfTableStyle style = CreateTableStyle(18);
            style.ClipTextToCellBounds = true;
            return RenderPaintTable(frame, options, new PdfTableCell(runs).WithNoWrap(), style);
        }
        Assert.Equal(Render(false), Render(true));
        Assert.True(report.HasLoss);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ParagraphPaintOutsideItsHorizontalFrameIsPreserved(bool column) {
        var report = new PdfConversionReport();
        PdfDocument document = PdfDocument.Create(CreateOptions(report));
        if (column) document.Row(row => row.PercentColumn(100, content => content.Paragraph(paragraph =>
            paragraph.BackgroundColor(PdfColor.Gray).Text("Preserved"))));
        else document.Paragraph(paragraph => paragraph.BackgroundColor(PdfColor.Gray).Text("Preserved"));
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Contains("Preserved", pdf.GetPage(1).Text);
        var warning = Assert.Single(report.Warnings, warning => warning.Code == "paragraph-paint-overflow");
        Assert.Equal(PdfLayoutDiagnosticKind.Overflow, warning.LayoutDiagnostic!.Kind);
        Assert.Equal(PdfConversionWarningSeverity.Information, warning.Severity);
        Assert.Equal(OfficeConversionLossKind.None, warning.LossKind);
        Assert.False(report.HasLoss);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void OrdinaryParagraphAndRowReportGlyphAndInlineOverhang(bool column, bool inline) {
        var report = new PdfConversionReport();
        var document = PdfDocument.Create(CreateOptions(report));
        var runs = inline
            ? new[] { PdfTextRun.Inline(new PdfInlineBox(5.004D, 8, borderColor: PdfColor.Black, borderWidth: 2)) }
            : new[] { PdfTextRun.Normal("T", fontSize: 9, font: PdfStandardFont.TimesItalic) };
        var style = new PdfParagraphStyle { RightIndent = 212D - 5.004D };
        if (column) document.Row(row => row.PercentColumn(100, content =>
            content.Paragraph(paragraph => paragraph.Runs(runs), style: style)));
        else document.Paragraph(paragraph => paragraph.Runs(runs), style: style);
        Assert.NotEmpty(document.ToBytes());
        var warning = Assert.Single(report.Warnings, warning => warning.Code == "paragraph-paint-overflow");
        Assert.Equal(PdfLayoutDiagnosticKind.Overflow, warning.LayoutDiagnostic!.Kind);
        Assert.Equal(OfficeConversionLossKind.None, warning.LossKind);
        Assert.False(report.HasLoss);
    }

    [Fact]
    public void EmptyInlineBoxAndWhitespaceHaveNoPaintToClip() {
        var report = new PdfConversionReport();
        byte[] bytes = PdfDocument.Create(CreateOptions(report)).Canvas(canvas => canvas.TextBox(new[] {
            PdfTextRun.Normal("   "), PdfTextRun.Inline(new PdfInlineBox(8, 40))
        }, 24, 24, 12, 4, new PdfCanvasTextBoxStyle { FontSize = 9, PaddingX = 0, PaddingY = 0 })).ToBytes();
        Assert.NotEmpty(bytes);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "canvas-text-box-paint-clipped");
    }

    [Fact]
    public void FilledPathHoleRemovesInkEvenWhenThePaintBoundsStayTheSame() {
        var report = new PdfConversionReport();
        var clip = OfficeClipPath.Path(new[] {
            OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(120, 0),
            OfficePathCommand.LineTo(120, 18), OfficePathCommand.LineTo(0, 18), OfficePathCommand.Close(),
            OfficePathCommand.MoveTo(20, 0), OfficePathCommand.LineTo(40, 0),
            OfficePathCommand.LineTo(40, 12), OfficePathCommand.LineTo(20, 12), OfficePathCommand.Close()
        }, OfficeFillRule.EvenOdd);
        byte[] bytes = PdfDocument.Create(CreateOptions(report)).Canvas(canvas => canvas.Clip(24, 24, clip,
            child => child.Text(new[] { PdfTextRun.Inline(new PdfInlineBox(120, 8, background: PdfColor.Black)) },
                24, 24, 120, 18))).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.NotEmpty(bytes);
        Assert.Single(report.Warnings, warning => warning.Code == "canvas-text-canvas-clip-paint-clipped");
        Assert.True(report.HasLoss);
    }

    [Fact]
    public void FrameDoesNotClaimLossOfPaintAlreadyRemovedByAnAncestor() {
        var report = new PdfConversionReport();
        byte[] bytes = PdfDocument.Create(CreateOptions(report)).Canvas(canvas => canvas.Clip(24, 24, 30, 18,
            child => child.Text(new[] { PdfTextRun.Inline(new PdfInlineBox(40, 8, background: PdfColor.Black, borderColor: PdfColor.Black, borderWidth: 2)) },
                24, 24, 40, 18))).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.NotEmpty(bytes);
        Assert.Single(report.Warnings, warning => warning.Code == "canvas-text-canvas-clip-paint-clipped");
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "canvas-text-paint-clipped");
    }

    private static byte[] RenderPaintTable(string frame, PdfCreateOptions options, PdfTableCell cell, PdfTableStyle style) {
        var document = PdfDocument.Create(options);
        var rows = new[] { new[] { cell } };
        return frame switch {
            "flow" => document.Table(rows, style: style).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }),
            "column" => document.Compose(compose => compose.Page(page => page.Content(content =>
                content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style)))))).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }),
            "canvas" => document.Canvas(canvas => canvas.Table(rows, 24, 24, 120, 18, style)).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }),
            _ => throw new ArgumentOutOfRangeException(nameof(frame))
        };
    }
}
