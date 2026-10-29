using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDrawingInteropTests {
    [Fact]
    public void OptionsClone_RetainsNormalizedColorsAlphaThemeRolesAndNullDefaults() {
        OfficeColor precise = OfficeColor.FromNormalizedRgb(0.12345, 0.23456, 0.34567, alpha: 77);
        var options = new PdfCreateOptions {
            DefaultTextColor = precise,
            HeaderTextColor = OfficeColor.Accent1,
            FooterTextColor = OfficeColor.Transparent,
            BackgroundColor = null,
            DefaultTableStyle = new PdfTableStyle { HeaderFill = precise }
        };

        PdfCreateOptions clone = options.Clone();
        options.DefaultTextColor = OfficeColor.Red;
        PdfTableStyle table = clone.DefaultTableStyle!;
        table.HeaderFill = OfficeColor.Blue;

        Assert.Equal(precise, clone.DefaultTextColor);
        Assert.Equal((0.12345, 0.23456, 0.34567), clone.DefaultTextColor!.Value.ToNormalizedRgb());
        Assert.Equal((byte)77, clone.DefaultTextColor.Value.A);
        Assert.Equal(OfficeColor.Accent1, clone.HeaderTextColor);
        Assert.Equal(OfficeColor.Transparent, clone.FooterTextColor);
        Assert.Null(clone.BackgroundColor);
        Assert.Equal(precise, clone.DefaultTableStyle!.HeaderFill);
    }

    [Fact]
    public void PdfDocument_WritesNormalizedRootColorWithoutByteProjection() {
        OfficeColor precise = OfficeColor.FromNormalizedRgb(0.12345, 0.23456, 0.34567, alpha: 77);
        var options = new PdfCreateOptions { DefaultTextColor = precise };
        using var document = PdfDocument.Create(options);
        document.Paragraph(p => p.Text("Precise normalized color"));
        options.DefaultTextColor = OfficeColor.Red;

        byte[] saved = document.ToBytes();
        string content = PdfTestContentText.ReadDecoded(saved);

        Assert.Contains("0.123 0.235 0.346 rg", content, StringComparison.Ordinal);
        Assert.DoesNotContain("0.122 0.235 0.345 rg", content, StringComparison.Ordinal);
        Assert.Single(PdfReadDocument.Open(saved).Pages);
    }

    [Fact]
    public void PdfDocument_AcceptsByteRgbOfficeColor() {
        var doc = PdfDocument.Create();
        doc.Paragraph(p => p.Color(OfficeColor.CornflowerBlue).Text("Drawing color reuse"));

        byte[] bytes = doc.ToBytes();

        string pdfContent = PdfTestContentText.ReadDecoded(bytes);
        Assert.Contains("0.392 0.584 0.929 rg", pdfContent);
    }

    [Fact]
    public void PdfDocument_SnapshotsSharedShapeDescriptorAtAddTime() {
        var shape = OfficeShape.Rectangle(90, 24);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.SteelBlue, OfficeColor.WhiteSmoke);
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.5;

        var doc = PdfDocument.Create(new PdfCreateOptions {
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

        string pdfContent = PdfTestContentText.ReadDecoded(bytes);
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
            KeepWithNext = true,
            ConstrainToContentWidth = true
        };
        var options = new PdfCreateOptions {
            DefaultDrawingStyle = style
        };

        style.Align = PdfAlign.Right;
        style.SpacingBefore = 1;
        style.SpacingAfter = 2;
        style.KeepWithNext = false;
        style.ConstrainToContentWidth = false;

        PdfDrawingStyle readback = options.DefaultDrawingStyle!;
        readback.Align = PdfAlign.Left;
        readback.SpacingAfter = 20;

        PdfCreateOptions clone = options.Clone();

        Assert.Equal(PdfAlign.Center, options.DefaultDrawingStyle!.Align);
        Assert.Equal(4, options.DefaultDrawingStyle.SpacingBefore);
        Assert.Equal(9, options.DefaultDrawingStyle.SpacingAfter);
        Assert.True(options.DefaultDrawingStyle.KeepWithNext);
        Assert.True(options.DefaultDrawingStyle.ConstrainToContentWidth);
        Assert.Equal(PdfAlign.Center, clone.DefaultDrawingStyle!.Align);
        Assert.Equal(9, clone.DefaultDrawingStyle.SpacingAfter);
        Assert.True(clone.DefaultDrawingStyle.KeepWithNext);
        Assert.True(clone.DefaultDrawingStyle.ConstrainToContentWidth);
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
        var options = new PdfCreateOptions().ApplyTheme(theme);

        drawingStyle.Align = PdfAlign.Right;
        drawingStyle.SpacingAfter = 1;
        drawingStyle.KeepWithNext = false;

        PdfCreateOptions clone = options.Clone();

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

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
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
                            row.PercentColumn(100, column =>
                                column.Shape(shape))))))
            .ToBytes();

        style.Align = PdfAlign.Right;
        style.SpacingBefore = 0;

        var (objects, _) = PdfSyntax.ParseObjects(bytes);
        string ReadPageProgram(PdfObject value) {
            if (value is PdfReference reference) return ReadPageProgram(objects[reference.ObjectNumber].Value);
            if (value is PdfStream stream) return PdfEncoding.Latin1GetString(
                OfficeIMO.Pdf.Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects));
            var parts = Assert.IsType<PdfArray>(value);
            var content = new System.Text.StringBuilder();
            foreach (PdfObject part in parts.Items) content.AppendLine(ReadPageProgram(part));
            return content.ToString();
        }

        int centeredPlacements = 0;
        foreach (PdfReadPage page in PdfReadDocument.Open(bytes).Pages) {
            var dictionary = Assert.IsType<PdfDictionary>(objects[page.ObjectNumber].Value);
            // Shared storage still executes once for each referencing page.
            string pageContent = ReadPageProgram(dictionary.Items["Contents"]);
            centeredPlacements += CountOccurrences(pageContent, "100 140 40 20 re f");
            Assert.DoesNotContain("20 140 40 20 re f", pageContent, StringComparison.Ordinal);
        }
        Assert.Equal(2, centeredPlacements);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PdfDocument_SnapshotsSharedDrawingSceneAtAddTime(bool scaledCanvas) {
        var background = OfficeShape.Rectangle(120, 60);
        background.FillColor = OfficeColor.WhiteSmoke;

        var drawing = new OfficeDrawing(120, 60)
            .AddShape(background, 0, 0);

        var doc = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            });
        if (scaledCanvas) doc.Canvas(canvas => canvas.Drawing(drawing, 30D, 40D, 180D, 90D));
        else doc.Drawing(drawing, align: PdfAlign.Center, spacingBefore: 4, spacingAfter: 6);

        drawing.Shapes[0].Shape.Width = 20;
        drawing.Shapes[0].Shape.FillColor = OfficeColor.Red;

        var laterShape = OfficeShape.Rectangle(10, 10);
        laterShape.FillColor = OfficeColor.Black;
        drawing.AddShape(laterShape, 0, 0);

        byte[] bytes = doc.ToBytes();

        string pdfContent = PdfTestContentText.ReadDecoded(bytes);
        Assert.Contains("0.961 0.961 0.961 rg", pdfContent, StringComparison.Ordinal);
        Assert.Contains(scaledCanvas ? "30 80 120 60 re f" : "60 100 120 60 re f", pdfContent, StringComparison.Ordinal);
        if (scaledCanvas) Assert.Contains("1.5 0 0 1.5", pdfContent, StringComparison.Ordinal);
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
