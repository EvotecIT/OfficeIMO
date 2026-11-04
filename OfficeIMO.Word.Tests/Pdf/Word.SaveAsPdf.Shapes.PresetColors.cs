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
    [InlineData("red", false, false, 255, 0, 0)]
    [InlineData("dkBlue", true, false, 0, 0, 139)]
    [InlineData("medVioletRed", false, true, 199, 21, 133)]
    public void SaveAsPdf_DrawingMlPresetColorTokensRetainFillAlphaAndText(
        string token, bool gradient, bool themeEntry, int red, int green, int blue) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Before preset shape");
        document.AddParagraph().AddShapeDrawing(WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(112D), OfficeIMO.OfficeLength.Points(48D));
        document.AddParagraph("After preset shape");
        var preset = new A.PresetColor(new A.Alpha { Val = 45000 });
        preset.SetAttribute(new OpenXmlAttribute("val", string.Empty, token));
        var main = document.MainDocumentPartRoot;
        OpenXmlElement color = preset;
        if (themeEntry) {
            A.Accent1Color accent = main.ThemePart!.Theme!.ThemeElements!.ColorScheme!.Accent1Color!;
            accent.RemoveAllChildren();
            accent.Append(preset);
            color = new A.SchemeColor { Val = A.SchemeColorValues.Accent1 };
        }
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

        byte[] bytes = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using (var pdf = PdfPigDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Before preset shape", text);
            Assert.Contains("After preset shape", text);
        }
        var primitives = PdfCore.PdfReadDocument.Open(bytes).Pages[0].GetIdentityVisualPrimitives();
        if (gradient) {
            var (start, end) = ReadDrawingGradientEndpoints(bytes, "DeviceRGB", 3);
            Assert.Equal(red, (int)System.Math.Round(start[0] * 255D));
            Assert.Equal(green, (int)System.Math.Round(start[1] * 255D));
            Assert.Equal(blue, (int)System.Math.Round(start[2] * 255D));
            foreach (double component in end) Assert.Equal(255, (int)System.Math.Round(component * 255D));
            var (startAlpha, endAlpha) = ReadDrawingGradientEndpoints(bytes, "DeviceGray", 1);
            Assert.InRange(startAlpha[0], 0.45D - 1D / 255D, 0.45D + 1D / 255D);
            Assert.Equal(1D, endAlpha[0], precision: 6);
            var paint = Assert.Single(primitives, primitive => primitive.FillGradient != null || primitive.FunctionPaint != null);
            Assert.Equal(1D, paint.FillOpacity ?? 1D, precision: 6);
        } else {
            var paint = Assert.Single(primitives, primitive => primitive.FillColor == OfficeColor.FromRgb((byte)red, (byte)green, (byte)blue));
            Assert.InRange(paint.FillOpacity ?? 1D, 0.45D - 1D / 255D, 0.45D + 1D / 255D);
        }
        var stroke = Assert.Single(primitives, primitive => primitive.StrokeColor == OfficeColor.FromRgb(51, 102, 153));
        Assert.Equal(112D, stroke.Width, precision: 3);
        Assert.Equal(48D, stroke.Height, precision: 3);
        Assert.Equal(1D, stroke.StrokeWidth, precision: 3);
        Assert.InRange(stroke.StrokeOpacity ?? 1D, 0.65D - 1D / 255D, 0.65D + 1D / 255D);
    }
}
