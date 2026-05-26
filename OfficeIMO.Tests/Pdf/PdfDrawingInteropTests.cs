using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDrawingInteropTests {
    [Fact]
    public void PdfColor_ConvertsToAndFromOfficeColor() {
        OfficeColor officeColor = OfficeColor.Parse("#336699CC");

        PdfColor pdfColor = PdfColor.FromOfficeColor(officeColor);
        OfficeColor roundTrip = pdfColor.ToOfficeColor(officeColor.A);

        Assert.Equal(0x33 / 255.0, pdfColor.R, 6);
        Assert.Equal(0x66 / 255.0, pdfColor.G, 6);
        Assert.Equal(0x99 / 255.0, pdfColor.B, 6);
        Assert.Equal(officeColor, roundTrip);
    }

    [Fact]
    public void PdfColor_FromOfficeColorOrNull_ReturnsNullForTransparentColor() {
        Assert.Null(PdfColor.FromOfficeColorOrNull(OfficeColor.Transparent));
        Assert.NotNull(PdfColor.FromOfficeColorOrNull(OfficeColor.FromRgba(10, 20, 30, 1)));
    }

    [Theory]
    [InlineData(-0.001, 0, 0, "r")]
    [InlineData(0, 1.001, 0, "g")]
    [InlineData(0, 0, double.NaN, "b")]
    [InlineData(double.PositiveInfinity, 0, 0, "r")]
    public void PdfColor_RejectsInvalidRgbComponents(double r, double g, double b, string paramName) {
        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfColor(r, g, b));

        Assert.Equal(paramName, exception.ParamName);
        Assert.Contains("PDF color components must be finite values between 0 and 1.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDoc_AcceptsOfficeColorThroughPdfColorConversion() {
        var doc = PdfDoc.Create();
        doc.Paragraph(p => p.Color(OfficeColor.CornflowerBlue).Text("Drawing color reuse"));

        byte[] bytes = doc.ToBytes();

        string pdfContent = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.392 0.584 0.929 rg", pdfContent);
    }

    [Fact]
    public void PdfDoc_SnapshotsSharedShapeDescriptorAtAddTime() {
        var shape = OfficeShape.Rectangle(90, 24);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.SteelBlue, OfficeColor.WhiteSmoke);
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.5;

        var doc = PdfDoc.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape);

        shape.Width = 10;
        shape.FillColor = OfficeColor.Red;
        shape.FillGradient = OfficeLinearGradient.Vertical(OfficeColor.Red, OfficeColor.Black);
        shape.StrokeColor = OfficeColor.Black;
        shape.StrokeWidth = 4;

        byte[] bytes = doc.ToBytes();

        string pdfContent = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/ShadingType 2", pdfContent, StringComparison.Ordinal);
        Assert.Contains("/C0 [0.275 0.51 0.706] /C1 [0.961 0.961 0.961]", pdfContent, StringComparison.Ordinal);
        Assert.Contains("0.275 0.51 0.706 RG", pdfContent, StringComparison.Ordinal);
        Assert.Contains("1.5 w", pdfContent, StringComparison.Ordinal);
        Assert.Contains("30 106 90 24 re W n", pdfContent, StringComparison.Ordinal);
        Assert.Contains("30 106 90 24 re S", pdfContent, StringComparison.Ordinal);
        Assert.DoesNotContain("1 0 0 rg", pdfContent, StringComparison.Ordinal);
        Assert.DoesNotContain("30 106 10 24 re", pdfContent, StringComparison.Ordinal);
    }

    [Fact]
    public void Options_SnapshotDefaultDrawingStyle() {
        var style = new PdfDrawingStyle {
            Align = PdfAlign.Center,
            SpacingBefore = 4,
            SpacingAfter = 9,
            KeepWithNext = true
        };
        var options = new PdfOptions {
            DefaultDrawingStyle = style
        };

        style.Align = PdfAlign.Right;
        style.SpacingBefore = 1;
        style.SpacingAfter = 2;
        style.KeepWithNext = false;

        PdfDrawingStyle readback = options.DefaultDrawingStyle!;
        readback.Align = PdfAlign.Left;
        readback.SpacingAfter = 20;

        PdfOptions clone = options.Clone();

        Assert.Equal(PdfAlign.Center, options.DefaultDrawingStyle!.Align);
        Assert.Equal(4, options.DefaultDrawingStyle.SpacingBefore);
        Assert.Equal(9, options.DefaultDrawingStyle.SpacingAfter);
        Assert.True(options.DefaultDrawingStyle.KeepWithNext);
        Assert.Equal(PdfAlign.Center, clone.DefaultDrawingStyle!.Align);
        Assert.Equal(9, clone.DefaultDrawingStyle.SpacingAfter);
        Assert.True(clone.DefaultDrawingStyle.KeepWithNext);
    }

    [Fact]
    public void Options_ApplyThemeSnapshotsDefaultDrawingStyle() {
        var drawingStyle = new PdfDrawingStyle {
            Align = PdfAlign.Center,
            SpacingAfter = 8,
            KeepWithNext = true
        };
        var theme = new PdfTheme {
            DrawingStyle = drawingStyle
        };
        var options = new PdfOptions().ApplyTheme(theme);

        drawingStyle.Align = PdfAlign.Right;
        drawingStyle.SpacingAfter = 1;
        drawingStyle.KeepWithNext = false;

        PdfOptions clone = options.Clone();

        Assert.Equal(PdfAlign.Center, options.DefaultDrawingStyle!.Align);
        Assert.Equal(8, options.DefaultDrawingStyle.SpacingAfter);
        Assert.True(options.DefaultDrawingStyle.KeepWithNext);
        Assert.Equal(PdfAlign.Center, clone.DefaultDrawingStyle!.Align);
        Assert.True(clone.DefaultDrawingStyle.KeepWithNext);
    }

    [Fact]
    public void DefaultDrawingStyle_AppliesAlignmentAndSpacingToFollowingShapesAndRows() {
        var style = new PdfDrawingStyle {
            Align = PdfAlign.Center,
            SpacingBefore = 4,
            SpacingAfter = 8
        };
        var shape = OfficeShape.Rectangle(40, 20);
        shape.FillColor = OfficeColor.WhiteSmoke;

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .DefaultDrawingStyle(style)
            .Shape(shape)
            .Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Shape(shape))))))
            .ToBytes();

        style.Align = PdfAlign.Right;
        style.SpacingBefore = 0;

        string pdfContent = Encoding.ASCII.GetString(bytes);
        Assert.True(CountOccurrences(pdfContent, "100 136 40 20 re f") >= 2, "Expected top-level and row-column shapes to inherit centered drawing placement and spacing.");
        Assert.DoesNotContain("20 140 40 20 re f", pdfContent, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDoc_SnapshotsSharedDrawingSceneAtAddTime() {
        var background = OfficeShape.Rectangle(120, 60);
        background.FillColor = OfficeColor.WhiteSmoke;

        var drawing = new OfficeDrawing(120, 60)
            .AddShape(background, 0, 0);

        var doc = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Drawing(drawing, align: PdfAlign.Center, spacingBefore: 4, spacingAfter: 6);

        drawing.Shapes[0].Shape.Width = 20;
        drawing.Shapes[0].Shape.FillColor = OfficeColor.Red;

        var laterShape = OfficeShape.Rectangle(10, 10);
        laterShape.FillColor = OfficeColor.Black;
        drawing.AddShape(laterShape, 0, 0);

        byte[] bytes = doc.ToBytes();

        string pdfContent = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.961 0.961 0.961 rg", pdfContent, StringComparison.Ordinal);
        Assert.Contains("60 96 120 60 re f", pdfContent, StringComparison.Ordinal);
        Assert.DoesNotContain("1 0 0 rg", pdfContent, StringComparison.Ordinal);
        Assert.DoesNotContain("0 0 0 rg", pdfContent, StringComparison.Ordinal);
        Assert.DoesNotContain("60 136 20 60 re f", pdfContent, StringComparison.Ordinal);
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += value.Length;
        }

        return count;
    }
}
