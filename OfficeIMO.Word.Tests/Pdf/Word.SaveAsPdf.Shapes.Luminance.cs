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
    [Fact]
    public void SaveAsPdf_DrawingMlLuminanceSequenceRetainsContinuousHslAcrossAlpha() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Before luminance shape");
        document.AddParagraph().AddShapeDrawing(WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(112D), OfficeIMO.OfficeLength.Points(48D));
        document.AddParagraph("After luminance shape");
        Wps.ShapeProperties properties = document.MainDocumentPartRoot.Document.Descendants<Wps.ShapeProperties>().Single();
        A.PresetGeometry? geometry = properties.GetFirstChild<A.PresetGeometry>();
        properties.RemoveAllChildren<A.NoFill>();
        properties.RemoveAllChildren<A.SolidFill>();
        properties.RemoveAllChildren<A.GradientFill>();
        var fill = new A.SolidFill(new A.RgbColorModelHex(
            new A.LuminanceModulation { Val = 50000 },
            new A.Alpha { Val = 45000 },
            new A.LuminanceOffset { Val = 20000 }) { Val = "4472C4" });
        if (geometry != null) properties.InsertAfter(fill, geometry);
        else properties.Append(fill);
        A.Outline outline = properties.GetFirstChild<A.Outline>() ?? new A.Outline();
        if (outline.Parent == null) properties.Append(outline);
        outline.Width = 12700;
        outline.RemoveAllChildren();
        outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "336699" }));

        byte[] bytes = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using (var pdf = PdfPigDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Before luminance shape", text);
            Assert.Contains("After luminance shape", text);
        }
        var primitives = PdfCore.PdfReadDocument.Open(bytes).Pages[0].GetIdentityVisualPrimitives();
        var paint = Assert.Single(primitives, primitive => primitive.FillColor == OfficeColor.FromRgb(56, 100, 178));
        Assert.InRange(paint.FillOpacity ?? 1D, 0.45D - 1D / 255D, 0.45D + 1D / 255D);
        Assert.Equal(112D, paint.Width, precision: 3);
        Assert.Equal(48D, paint.Height, precision: 3);
        Assert.Equal(OfficeColor.FromRgb(51, 102, 153), paint.StrokeColor);
        Assert.Equal(1D, paint.StrokeWidth, precision: 3);
        Assert.Equal(1D, paint.StrokeOpacity ?? 1D, precision: 6);
    }
}
