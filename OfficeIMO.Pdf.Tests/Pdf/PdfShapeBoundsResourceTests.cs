using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfShapeBoundsResourceTests {
    [Fact]
    public void SvgPreparationReturnsAReasonWhenStrokeProjectionExceedsItsLimit() {
        byte[] svg = System.Text.Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='1000' height='8'>" +
            "<path d='M0 4H1000' fill='none' stroke='blue' stroke-width='20000' stroke-dasharray='.25 .25' stroke-linecap='round'/></svg>");
        Assert.True(OfficeSvgDrawingReader.TryRead(svg, out _, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.False(PdfDocument.TryPrepareImageBytes(svg, out _, out _, out _, out string? reason));
        Assert.Contains("point limit", reason);
    }

    [Fact]
    public void DefaultStrokeProjectionAlsoBoundsDirectPainterExpansion() {
        var contours = new[] { new OfficeFlattenedPathContour(new[] { new OfficePoint(0, 0), new OfficePoint(200, 0) }, false) };
        var error = Assert.Throws<NotSupportedException>(() => OfficeStrokeGeometry.Create(contours, 20000,
            OfficeStrokeLineCap.Round, OfficeStrokeLineJoin.Round, 4, new[] { .0625, .0625 }, 0, 4,
            double.NegativeInfinity, double.NegativeInfinity, double.PositiveInfinity, double.PositiveInfinity));
        Assert.Contains("point limit", error.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RunningStrokeProjectionLimitKeepsTheDocumentAndReportsUnassessedBounds(bool footer) {
        // A bounded version of the reported 6000-point line; the full payload is
        // qualified separately without asking a broken engine to allocate gigabytes.
        OfficeShape shape = OfficeShape.Line(0, 0, 60, 0);
        shape.StrokeColor = OfficeColor.Blue;
        shape.StrokeWidth = 20000;
        shape.StrokeLineCap = OfficeStrokeLineCap.Round;
        shape.SetStrokeDashArray(new[] { .0625, .0625 });
        var report = new PdfConversionReport();
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = 200, PageHeight = 200, MarginLeft = 10, MarginRight = 10,
            MarginTop = 30, MarginBottom = 30, HeaderOffsetY = 0, FooterOffsetY = 0
        }.ReportDiagnosticsTo(report)).Compose(c => c.Page(page => {
            if (footer) page.Footer(builder => builder.Shape(shape));
            else page.Header(builder => builder.Shape(shape));
            page.Content(builder => builder.Paragraph(paragraph => paragraph.Text("MARKER")));
        })).ToBytes(new PdfSaveOptions { AllowLoss = true });

        Assert.Contains("MARKER", PdfReadDocument.Open(bytes).ExtractText());
        var warning = Assert.Single(report.Warnings, warning => warning.Code == "header-footer-paint-bounds-unmeasured");
        Assert.Equal(OfficeConversionLossKind.Unassessed, warning.LossKind);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "header-footer-page-bounds-clipped");
    }

    [Fact]
    public void LinkedStrokeProjectionLimitRetainsConservativeClickableBounds() {
        OfficeShape shape = OfficeShape.Line(0, 0, .9, 0);
        shape.StrokeColor = OfficeColor.Blue;
        shape.StrokeWidth = 1000;
        shape.StrokeLineCap = OfficeStrokeLineCap.Round;
        shape.Transform = OfficeTransform.Scale(100, 100);
        shape.SetStrokeDashArray(new[] { 1D / 1024, 1D / 1024 });
        byte[] bytes = PdfDocument.Create().Compose(c => c.Content(body => body.Shape(shape, linkUri: "https://example.test/shape"))).ToBytes();

        var annotation = Assert.Single(PdfDocument.Load(bytes).Inspect().LinkAnnotations);
        Assert.True(annotation.X2 - annotation.X1 >= 100000);
        Assert.True(annotation.Y2 - annotation.Y1 >= 100000);
    }

    [Fact]
    public void StrokeProjectionBudgetRejectsTheSourceBeforeCopyingItsPoints() {
        int remaining = 64;
        void Charge(int points) {
            if (points > remaining) throw new InvalidOperationException("Projection budget exhausted.");
            remaining -= points;
        }
        var contours = new[] { new OfficeFlattenedPathContour(new CountOnlyPoints(), false) };
        var error = Assert.Throws<InvalidOperationException>(() => OfficeStrokeGeometry.Create(contours, 2,
            OfficeStrokeLineCap.Round, OfficeStrokeLineJoin.Round, 4, null, 0, 4,
            double.NegativeInfinity, double.NegativeInfinity, double.PositiveInfinity, double.PositiveInfinity,
            chargePoints: Charge));
        Assert.Equal("Projection budget exhausted.", error.Message);
    }

    [Fact]
    public void StrokeProjectionBudgetPreservesAdmittedOutlineGeometry() {
        var contours = new[] { new OfficeFlattenedPathContour(new[] {
            new OfficePoint(0, 0), new OfficePoint(4, 0), new OfficePoint(4, 4)
        }, false) };
        int remaining = 512;
        void Charge(int points) {
            if (points > remaining) throw new InvalidOperationException("Projection budget exhausted.");
            remaining -= points;
        }
        List<List<OfficePoint>> Outline(Action<int>? charge) => OfficeStrokeGeometry.Create(contours, 2,
            OfficeStrokeLineCap.Round, OfficeStrokeLineJoin.Round, 4, new[] { 2D, 1D }, 0, 4,
            double.NegativeInfinity, double.NegativeInfinity, double.PositiveInfinity, double.PositiveInfinity,
            chargePoints: charge);
        var expected = Outline(null);
        var actual = Outline(Charge);
        Assert.NotEmpty(actual);
        Assert.True(remaining < 512);
        Assert.Equal(expected.SelectMany(ring => ring).ToArray(), actual.SelectMany(ring => ring).ToArray());
    }

    private sealed class CountOnlyPoints : IReadOnlyList<OfficePoint> {
        public int Count => 65;
        public OfficePoint this[int index] => throw new NotSupportedException("Source points must not be copied after admission fails.");
        public IEnumerator<OfficePoint> GetEnumerator() => throw new NotSupportedException("Source points must not be copied after admission fails.");
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
