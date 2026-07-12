using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageCropTranslateTests {
    [Fact]
    public void CropAndTranslate_ClipsContentAndTranslatesAnnotationGeometry() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Crop translation keeps source content"))
            .ToBytes();
        source = PdfAnnotationEditor.AddStampAnnotation(
            source,
            new PdfStampAnnotationOptions {
                X = 100,
                Y = 150,
                Width = 80,
                Height = 30,
                Contents = "Inside crop"
            }).Bytes;
        source = PdfAnnotationEditor.AddStampAnnotation(
            source,
            new PdfStampAnnotationOptions {
                X = 5,
                Y = 5,
                Width = 20,
                Height = 20,
                Contents = "Outside crop"
            }).Bytes;

        PdfDocument cropped = PdfDocument.Open(source)
            .Pages.CropAndTranslate(50, 50, 400, 700);
        byte[] output = cropped.ToBytes();
        PdfDocumentInfo info = cropped.Inspect();

        Assert.Equal(350, info.Pages[0].MediaBox!.Width);
        Assert.Equal(650, info.Pages[0].MediaBox!.Height);
        Assert.Equal(0, info.Pages[0].MediaBox!.Left);
        Assert.Equal(0, info.Pages[0].MediaBox!.Bottom);
        Assert.Equal(350, info.Pages[0].CropBox!.Width);
        Assert.Equal(650, info.Pages[0].CropBox!.Height);
        PdfAnnotation stamp = Assert.Single(info.GetAnnotationsBySubtype("Stamp"));
        Assert.Equal("Inside crop", stamp.Contents);
        Assert.Equal(50, stamp.X1);
        Assert.Equal(100, stamp.Y1);
        Assert.Equal(130, stamp.X2);
        Assert.Equal(130, stamp.Y2);
        Assert.Contains("Crop translation keeps source content", PdfTextExtractor.ExtractAllText(output), StringComparison.Ordinal);
        Assert.Contains("1 0 0 1 -50 -50 cm", PdfEncoding.Latin1GetString(output), StringComparison.Ordinal);
    }

    [Fact]
    public void CropAndTranslate_RejectsRectanglesOutsideTheMediaBox() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Invalid crop"))
            .ToBytes();

        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfPageEditor.CropAndTranslatePages(source, -1, 0, 100, 100));

        Assert.Contains("MediaBox", exception.Message, StringComparison.Ordinal);
    }
}
