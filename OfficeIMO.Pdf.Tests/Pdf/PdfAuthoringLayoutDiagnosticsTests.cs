using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfAuthoringLayoutDiagnosticsTests {
    [Theory]
    [InlineData("flow", true)]
    [InlineData("column", true)]
    [InlineData("canvas", true)]
    [InlineData("flow", false)]
    [InlineData("column", false)]
    [InlineData("canvas", false)]
    public void FixedHeightCellReportsContentLossOnlyWhenTextDoesNotFit(string frame, bool overflows) {
        var report = new PdfConversionReport();
        var cell = new PdfTableCell(overflows ? "VisibleLine\nClippedLine\nFullyClippedLine" : "VisibleLine");
        byte[] bytes = RenderTable(frame, report, cell, 18);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("VisibleLine", pdf.GetPage(1).Text);
        if (!overflows) {
            Assert.DoesNotContain(report.Warnings, warning => warning.Code == "table-cell-content-clipped");
            Assert.False(report.HasLoss);
            return;
        }

        if (frame != "canvas") Assert.DoesNotContain("ClippedLine", pdf.GetPage(1).Text);
        PdfConversionWarning warning = Assert.Single(report.Warnings, warning => warning.Code == "table-cell-content-clipped");
        Assert.Equal("AuthoringProbe", warning.Converter);
        Assert.Equal(PdfConversionWarningSeverity.Warning, warning.Severity);
        Assert.Equal(PdfLayoutDiagnosticKind.ClippedContent, warning.LayoutDiagnostic!.Kind);
        Assert.Equal(24D, warning.LayoutDiagnostic.X);
        Assert.Equal(24D, warning.LayoutDiagnostic.Y);
        Assert.Equal(120D, warning.LayoutDiagnostic.Width);
        Assert.Equal(18D, warning.LayoutDiagnostic.Height);
        Assert.True(report.HasLoss);
    }

    [Theory]
    [InlineData("flow", true)]
    [InlineData("column", true)]
    [InlineData("canvas", true)]
    [InlineData("flow", false)]
    [InlineData("column", false)]
    [InlineData("canvas", false)]
    public void CellObjectOverflowIsReportedAsPreservedContent(string frame, bool overflows) {
        var report = new PdfConversionReport();
        var cell = PdfTableCell.WithCheckBoxes(string.Empty, new[] {
            new PdfTableCellCheckBox("Probe", size: overflows ? 36 : 12)
        });
        byte[] bytes = RenderTable(frame, report, cell, 18);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(bytes).FormFields);
        Assert.Equal(overflows ? 36D : 12D, Assert.Single(field.Widgets).Height, 3);
        if (overflows) {
            PdfConversionWarning warning = Assert.Single(report.Warnings, warning => warning.Code == "table-cell-object-overflow");
            Assert.Equal(PdfConversionWarningSeverity.Information, warning.Severity);
            Assert.Equal(PdfLayoutDiagnosticKind.Overflow, warning.LayoutDiagnostic!.Kind);
            Assert.Equal(OfficeConversionLossKind.None, warning.LossKind);
        } else {
            Assert.DoesNotContain(report.Warnings, warning => warning.Code == "table-cell-object-overflow");
        }
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "table-cell-content-clipped");
        Assert.False(report.HasLoss);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TextBoxClippingReachesConversionReportWithoutRequiringCallback(bool overflows) {
        var report = new PdfConversionReport();
        byte[] bytes = PdfDocument.Create(CreateOptions(report)).Canvas(canvas => canvas.TextBox(
            overflows ? "First\nSecond\nThird" : "First", 24, 24, 120, 18,
            new PdfCanvasTextBoxStyle {
                FontSize = 9, PaddingX = 0, PaddingY = 0, Background = null, BorderColor = null
            })).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("First", pdf.GetPage(1).Text);
        if (overflows) {
            PdfConversionWarning warning = Assert.Single(report.Warnings, warning => warning.Code == "canvas-text-box-content-clipped");
            Assert.Equal(PdfLayoutDiagnosticKind.ClippedContent, warning.LayoutDiagnostic!.Kind);
            Assert.Equal(24D, warning.LayoutDiagnostic.Y);
            Assert.True(report.HasLoss);
        } else {
            Assert.DoesNotContain(report.Warnings, warning => warning.Code == "canvas-text-box-content-clipped");
            Assert.False(report.HasLoss);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PlainCanvasTextReportsFullyClippedLines(bool overflows) {
        var report = new PdfConversionReport();
        byte[] bytes = PdfDocument.Create(CreateOptions(report)).Canvas(canvas => canvas.Text(
            overflows ? "First\nSecond\nThird" : "First", 24, 24, 120, 18, fontSize: 9)).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("First", pdf.GetPage(1).Text);
        Assert.Equal(overflows, report.Warnings.Any(warning => warning.Code == "canvas-text-content-clipped"));
        Assert.Equal(overflows, report.HasLoss);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CanvasFramesDoNotReportUnusedLineAdvanceAsClippedPaint(bool table) {
        var report = new PdfConversionReport();
        PdfDocument document = PdfDocument.Create(CreateOptions(report));
        if (table) {
            PdfTableStyle style = CreateTableStyle(18);
            style.LineHeight = 12; // A 108-point advance; the sole 9-point glyph line still fits.
            document.Canvas(canvas => canvas.Table(new[] { new[] { "Fits" } }, 24, 24, 120, 18, style));
        } else {
            document.Canvas(canvas => canvas.TextBox("Fits", 24, 24, 120, 18,
                new PdfCanvasTextBoxStyle { FontSize = 9, LineHeight = 108, PaddingX = 0, PaddingY = 0 }));
        }
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Contains("Fits", pdf.GetPage(1).Text);
        Assert.DoesNotContain(report.Warnings, warning => warning.LayoutDiagnostic?.Kind == PdfLayoutDiagnosticKind.ClippedContent);
        Assert.False(report.HasLoss);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    [InlineData("canvas")]
    public void CellObjectBeyondPhysicalPageIsReportedAsClipped(string frame) {
        var report = new PdfConversionReport();
        var cell = PdfTableCell.WithCheckBoxes(string.Empty, new[] {
            new PdfTableCellCheckBox("OffPageProbe", size: 180)
        });
        PdfFormField field = Assert.Single(PdfInspector.Inspect(RenderTable(frame, report, cell, 18)).FormFields);
        Assert.True(Assert.Single(field.Widgets).Y1 < 0);
        PdfConversionWarning warning = Assert.Single(report.Warnings, warning => warning.Code == "table-cell-object-page-bounds-clipped");
        Assert.Equal(PdfLayoutDiagnosticKind.ClippedContent, warning.LayoutDiagnostic!.Kind);
        Assert.True(report.HasLoss);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "table-cell-object-overflow");
    }

    [Theory]
    [InlineData("flow", true)]
    [InlineData("column", true)]
    [InlineData("canvas", true)]
    [InlineData("flow", false)]
    [InlineData("column", false)]
    [InlineData("canvas", false)]
    public void NoWrapCellReportsHorizontalClipping(string frame, bool overflows) {
        var report = new PdfConversionReport();
        var cell = new PdfTableCell(overflows ? new string('W', 128) : "Fits").WithNoWrap();
        using PdfPigDocument pdf = PdfPigDocument.Open(RenderTable(frame, report, cell, 18));
        Assert.Equal(1, pdf.NumberOfPages);
        if (overflows) {
            PdfConversionWarning warning = Assert.Single(report.Warnings, warning => warning.Code == "table-cell-content-clipped");
            Assert.Equal(PdfLayoutDiagnosticKind.ClippedContent, warning.LayoutDiagnostic!.Kind);
            Assert.Equal(120D, warning.LayoutDiagnostic.Width);
            Assert.True(report.HasLoss);
        } else {
            Assert.Contains("Fits", pdf.GetPage(1).Text);
            Assert.DoesNotContain(report.Warnings, warning => warning.Code == "table-cell-content-clipped");
            Assert.False(report.HasLoss);
        }
    }

    [Theory]
    [InlineData("flow", false)]
    [InlineData("column", false)]
    [InlineData("canvas", false)]
    [InlineData("flow", true)]
    [InlineData("column", true)]
    [InlineData("canvas", true)]
    public void NoWrapDiagnosticsUseTheFinalShrunkTextGeometry(string frame, bool explicitFontSize) {
        var report = new PdfConversionReport();
        const int letters = 20;
        var text = new string('W', letters);
        var cell = (explicitFontSize
            ? new PdfTableCell(new[] { new PdfTextRun(text, fontSize: 14) })
            : new PdfTableCell(text)).WithNoWrap().WithShrinkTextToFit();
        using PdfPigDocument pdf = PdfPigDocument.Open(RenderTable(frame, report, cell, 18));
        var painted = pdf.GetPage(1).Letters;
        Assert.Equal(letters, painted.Count);
        Assert.All(painted, letter => {
            Assert.InRange(letter.PointSize, 6D, (explicitFontSize ? 14D : 9D) - .01D);
            Assert.True(letter.BoundingBox.Right <= 144.01D);
        });
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "table-cell-content-clipped");
        Assert.False(report.HasLoss);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    public void ContinuingTableRowDoesNotReportClipping(string frame) {
        var report = new PdfConversionReport();
        var cell = new PdfTableCell(string.Join("\n", Enumerable.Range(1, 25).Select(index => "Line" + index)));
        using PdfPigDocument pdf = PdfPigDocument.Open(RenderTable(frame, report, cell, null));
        Assert.True(pdf.NumberOfPages > 1);
        Assert.Contains("Line25", pdf.GetPage(pdf.NumberOfPages).Text);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "table-cell-content-clipped");
        Assert.False(report.HasLoss);
    }

    [Fact]
    public void SeparateClippedCellsKeepDistinctOccurrences() {
        var report = new PdfConversionReport();
        var style = CreateTableStyle(18);
        style.FixedRowHeights = new List<double?> { 18, 18 };
        byte[] bytes = PdfDocument.Create(CreateOptions(report)).Table(new[] {
            new[] { new PdfTableCell("First\nHiddenOne") },
            new[] { new PdfTableCell("Second\nHiddenTwo") }
        }, style: style).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("First", pdf.GetPage(1).Text);
        Assert.Contains("Second", pdf.GetPage(1).Text);
        PdfConversionWarning[] warnings = report.Warnings.Where(warning => warning.Code == "table-cell-content-clipped").ToArray();
        Assert.Equal(2, warnings.Length);
        Assert.NotEqual(warnings[0].Source, warnings[1].Source);
    }

    private static byte[] RenderTable(string frame, PdfConversionReport report, PdfTableCell cell, double? fixedHeight) {
        PdfDocument document = PdfDocument.Create(CreateOptions(report));
        var rows = new[] { new[] { cell } };
        PdfTableStyle style = CreateTableStyle(fixedHeight);
        return frame switch {
            "flow" => document.Table(rows, style: style).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }),
            "column" => document.Compose(compose => compose.Page(page => page.Content(content =>
                content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style)))))).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }),
            "canvas" => document.Canvas(canvas => canvas.Table(rows, 24, 24, 120, fixedHeight!.Value, style)).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }),
            _ => throw new ArgumentOutOfRangeException(nameof(frame))
        };
    }

    private static PdfCreateOptions CreateOptions(PdfConversionReport report) => new PdfCreateOptions {
        PageWidth = global::OfficeIMO.OfficeLength.Points(260), PageHeight = global::OfficeIMO.OfficeLength.Points(180),
        MarginLeft = global::OfficeIMO.OfficeLength.Points(24), MarginRight = global::OfficeIMO.OfficeLength.Points(24), MarginTop = global::OfficeIMO.OfficeLength.Points(24), MarginBottom = global::OfficeIMO.OfficeLength.Points(24),
        DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(9)
    }.ReportDiagnosticsTo(report, "AuthoringProbe");

    private static PdfTableStyle CreateTableStyle(double? fixedHeight) {
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.RepeatHeaderRowCount = 0;
        style.FontSize = 9;
        style.LineHeight = 4D / 3D; // Table line height is a multiplier: 9-point text advances 12 points.
        style.CellPaddingX = 0;
        style.CellPaddingY = 0;
        style.CellSpacing = 0;
        style.SpacingBefore = 0;
        style.SpacingAfter = 0;
        style.ColumnWidthPoints = new List<double?> { 120 };
        style.FixedRowHeights = fixedHeight.HasValue ? new List<double?> { fixedHeight } : null;
        return style;
    }
}
