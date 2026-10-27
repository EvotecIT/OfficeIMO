using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("srgb", "alpha-offset", 0.5D)]
    [InlineData("scheme", "alpha-offset", 0.5D)]
    [InlineData("preset", "alpha-offset", 0.5D)]
    [InlineData("srgb", "offset-modulate", 0.4D)]
    [InlineData("srgb", "modulate-offset", 0.3D)]
    [InlineData("srgb", "modulate-chain", 0.2D)]
    [InlineData("srgb", "upper-clamp", 0.5D)]
    [InlineData("srgb", "lower-clamp", 0.2D)]
    [InlineData("theme-alpha", "modulate-offset", 0.1D)]
    [InlineData("gradient", "alpha-offset", 0.5D)]
    [InlineData("gradient", "modulate-offset", 0.3D)]
    public void SaveAsPdf_DrawingMlFillAndOutlineShareOrderedAlphaTransforms(string colorKind, string transforms, double expected) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph().AddShapeDrawing(WordShapeType.Rectangle, 112D, 48D);
        OpenXmlElement color;
        if (colorKind == "scheme" || colorKind == "theme-alpha") {
            var accent = document.MainDocumentPartRoot.ThemePart!.Theme!.ThemeElements!.ColorScheme!.Accent1Color!;
            accent.RemoveAllChildren();
            var themeColor = new A.RgbColorModelHex { Val = "336699" };
            if (colorKind == "theme-alpha") themeColor.Append(new A.Alpha { Val = 60000 });
            accent.Append(themeColor);
            color = new A.SchemeColor { Val = A.SchemeColorValues.Accent1 };
        } else if (colorKind == "preset") {
            color = new A.PresetColor { Val = A.PresetColorValues.Red };
        } else {
            color = new A.RgbColorModelHex { Val = "336699" };
        }
        switch (transforms) {
            case "alpha-offset":
                color.Append(new A.Alpha { Val = 65000 }, new A.AlphaOffset { Val = -15000 });
                break;
            case "offset-modulate":
                color.Append(new A.AlphaOffset { Val = -20000 }, new A.AlphaModulation { Val = 50000 });
                break;
            case "modulate-offset":
                color.Append(new A.AlphaModulation { Val = 50000 }, new A.AlphaOffset { Val = -20000 });
                break;
            case "modulate-chain":
                color.Append(new A.Alpha { Val = 80000 }, new A.AlphaModulation { Val = 50000 }, new A.AlphaModulation { Val = 50000 });
                break;
            case "upper-clamp":
                color.Append(new A.AlphaOffset { Val = 20000 }, new A.AlphaModulation { Val = 50000 });
                break;
            case "lower-clamp":
                color.Append(new A.AlphaOffset { Val = -100000 }, new A.AlphaOffset { Val = 20000 });
                break;
        }
        var properties = document.MainDocumentPartRoot.Document.Descendants<Wps.ShapeProperties>().Single();
        properties.RemoveAllChildren<A.NoFill>();
        properties.RemoveAllChildren<A.SolidFill>();
        if (colorKind == "gradient") {
            properties.AddChild(new A.GradientFill(new A.GradientStopList(
                new A.GradientStop(color.CloneNode(true)) { Position = 0 },
                new A.GradientStop(new A.RgbColorModelHex { Val = "FFFFFF" }) { Position = 100000 }),
                new A.LinearGradientFill { Angle = 0 }), true);
        } else {
            properties.AddChild(new A.SolidFill(color.CloneNode(true)), true);
        }
        properties.RemoveAllChildren<A.Outline>();
        properties.AddChild(new A.Outline(new A.SolidFill(color)) { Width = 50800 }, true);

        byte[] bytes = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        var page = PdfCore.PdfReadDocument.Open(bytes).Pages[0];
        var paint = Assert.Single(page.GetIdentityVisualPrimitives(), primitive => primitive.StrokeWidth == 4D);
        if (colorKind == "gradient") {
            var (start, end) = ReadDrawingGradientEndpoints(bytes, "DeviceGray", 1);
            Assert.InRange(start[0], expected - 2D / 255D, expected + 2D / 255D);
            Assert.Equal(1D, end[0], precision: 6);
            Assert.Equal(1D, paint.FillOpacity ?? 1D, precision: 6);
        } else {
            Assert.InRange(paint.FillOpacity ?? 1D, expected - 2D / 255D, expected + 2D / 255D);
            Assert.Equal(paint.FillOpacity, paint.StrokeOpacity);
        }
        Assert.InRange(paint.StrokeOpacity ?? 1D, expected - 2D / 255D, expected + 2D / 255D);

        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_WORD_ALPHA_PROOF_OUTPUT");
        if (!string.IsNullOrWhiteSpace(output)) {
            Directory.CreateDirectory(output);
            string stem = Path.Combine(output, colorKind + "-" + transforms);
            File.WriteAllBytes(stem + ".pdf", bytes);
            File.WriteAllBytes(stem + ".png", document.ToPng(new WordImageExportOptions { Scale = 1D }));
        }
    }
}
