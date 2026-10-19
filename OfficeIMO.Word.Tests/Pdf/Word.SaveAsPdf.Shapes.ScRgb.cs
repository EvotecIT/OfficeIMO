using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false, false, 188)]
    [InlineData(false, false, true, 231)]
    [InlineData(true, false, true, 231)]
    [InlineData(false, true, true, 231)]
    [InlineData(true, true, true, 231)]
    public void SaveAsPdf_DrawingMlScRgbUsesSharedLinearColorDecoding(bool gradient, bool themeEntry, bool tinted, int expectedChannel) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Before scRGB shape");
        document.AddParagraph().AddShapeDrawing(WordShapeType.Rectangle, 112D, 48D);
        document.AddParagraph("After scRGB shape");
        var main = document.MainDocumentPartRoot;
        if (themeEntry) {
            A.Accent1Color accent = main.ThemePart!.Theme!.ThemeElements!.ColorScheme!.Accent1Color!;
            accent.RemoveAllChildren();
            accent.Append(CreateNativeScRgbColor());
        }
        OpenXmlElement color = themeEntry
            ? new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }
            : CreateNativeScRgbColor();
        color.Append(new A.Alpha { Val = 45000 });
        if (tinted) color.Append(new A.Tint { Val = 40000 });
        Wps.ShapeProperties properties = main.Document.Descendants<Wps.ShapeProperties>().Single();
        A.PresetGeometry? geometry = properties.GetFirstChild<A.PresetGeometry>();
        properties.RemoveAllChildren<A.NoFill>();
        properties.RemoveAllChildren<A.SolidFill>();
        properties.RemoveAllChildren<A.GradientFill>();
        OpenXmlElement fill = gradient
            ? new A.GradientFill(new A.GradientStopList(
                new A.GradientStop(color) { Position = 0 },
                new A.GradientStop(new A.RgbColorModelHex { Val = "FFFFFF" }) { Position = 100000 }),
                new A.LinearGradientFill { Angle = 0 })
            : new A.SolidFill(color);
        if (geometry != null) properties.InsertAfter(fill, geometry);
        else properties.Append(fill);
        A.Outline outline = properties.GetFirstChild<A.Outline>() ?? new A.Outline();
        if (outline.Parent == null) properties.Append(outline);
        outline.Width = 12700;
        outline.RemoveAllChildren();
        outline.Append(new A.SolidFill(new A.RgbColorModelHex(new A.Alpha { Val = 65000 }) { Val = "336699" }));

        byte[] bytes = document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false });
        using (var pdf = PdfPigDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Before scRGB shape", text);
            Assert.Contains("After scRGB shape", text);
        }
        var primitives = PdfCore.PdfReadDocument.Open(bytes).Pages[0].GetIdentityVisualPrimitives();
        if (gradient) {
            var (start, end) = ReadDrawingGradientEndpoints(bytes, "DeviceRGB", 3);
            for (int component = 0; component < 3; component++) {
                Assert.Equal(expectedChannel, (int)System.Math.Round(start[component] * 255D));
                Assert.Equal(255, (int)System.Math.Round(end[component] * 255D));
            }
            var (startAlpha, endAlpha) = ReadDrawingGradientEndpoints(bytes, "DeviceGray", 1);
            Assert.InRange(startAlpha[0], 0.45D - 1D / 255D, 0.45D + 1D / 255D);
            Assert.Equal(1D, endAlpha[0], precision: 6);
            var paint = Assert.Single(primitives, primitive => primitive.FillGradient != null || primitive.FunctionPaint != null);
            Assert.Equal(1D, paint.FillOpacity ?? 1D, precision: 6);
        } else {
            var paint = Assert.Single(primitives, primitive => primitive.FillColor == OfficeColor.FromRgb(
                (byte)expectedChannel, (byte)expectedChannel, (byte)expectedChannel));
            Assert.InRange(paint.FillOpacity ?? 1D, 0.45D - 1D / 255D, 0.45D + 1D / 255D);
        }
        var stroke = Assert.Single(primitives, primitive => primitive.StrokeColor == OfficeColor.FromRgb(51, 102, 153));
        Assert.Equal(112D, stroke.Width, precision: 3);
        Assert.Equal(48D, stroke.Height, precision: 3);
        Assert.Equal(1D, stroke.StrokeWidth, precision: 3);
        Assert.Equal(0.65D, stroke.StrokeOpacity ?? 1D, precision: 3);
    }

    private static A.RgbColorModelPercentage CreateNativeScRgbColor() => new() {
        RedPortion = 50000, GreenPortion = 50000, BluePortion = 50000
    };
}
