using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfImageOrientationPlacementTests {
    public static IEnumerable<object[]> DrawingOrientations() {
        for (int orientation = 1; orientation <= 8; orientation++) {
            yield return new object[] { orientation, false };
            yield return new object[] { orientation, true };
        }
    }

    [Theory]
    [MemberData(nameof(DrawingOrientations))]
    public void ManagedDisplayAndDrawingReauthoringApplyExifExactlyOnce(int orientation, bool clip) {
        byte[] jpeg = CreateJpeg(orientation);
        Assert.True(OfficeRasterImageDecoder.TryDecode(jpeg, out OfficeRasterImage? reference));
        int width = reference!.Width * 20, height = reference.Height * 20;
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(width), PageHeight = global::OfficeIMO.OfficeLength.Points(height),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(0), MarginRight = global::OfficeIMO.OfficeLength.Points(0), MarginTop = global::OfficeIMO.OfficeLength.Points(0), MarginBottom = global::OfficeIMO.OfficeLength.Points(0) };
        byte[] pdf = PdfDocument.Create(builder => builder.Content(content => content.Image(jpeg, global::OfficeIMO.OfficeLength.Points(width), global::OfficeIMO.OfficeLength.Points(height),
            style: new PdfImageStyle { ClipPath = clip ? OfficeClipPath.RoundedRectangle(width, height, Math.Min(width, height) / 4D) : null })), options).ToBytes();
        var document = PdfDocument.Load(pdf);
        PdfExtractedImage extracted = Assert.Single(document.Reader.Images());
        Assert.Equal(jpeg, extracted.Bytes);
        CheckPixels(pdf);

        OfficeDrawing drawing = document.Render.Drawing(1);
        byte[] replay = PdfDocument.Create(builder => builder.Content(content =>
            content.Drawing(drawing, spacingBefore: global::OfficeIMO.OfficeLength.Points(0), spacingAfter: global::OfficeIMO.OfficeLength.Points(0))), options).ToBytes();
        CheckPixels(replay);
        Assert.Equal(jpeg, extracted.Bytes); // Presentation must not mutate the extraction contract.
        Assert.Equal(jpeg, Assert.Single(document.Reader.Images()).Bytes);

        void CheckPixels(byte[] bytes) {
            PdfPageRenderResult rendered = Assert.Single(PdfDocument.Load(bytes).Render.Pages(options: new PdfPageRenderOptions { Dpi = 72 }));
            Assert.NotNull(rendered.Bytes);
            Assert.True(OfficeRasterImageDecoder.TryDecode(rendered.Bytes!, out OfficeRasterImage? actual));
            Assert.Equal(width, actual!.Width);
            Assert.Equal(height, actual.Height);
            for (int y = 0; y < reference.Height; y++) for (int x = 0; x < reference.Width; x++) {
                OfficeColor expected = reference.GetPixel(x, y), found = actual.GetPixel(x * 20 + 10, y * 20 + 10);
                Assert.InRange(Math.Abs(found.R - expected.R), 0, 3);
                Assert.InRange(Math.Abs(found.G - expected.G), 0, 3);
                Assert.InRange(Math.Abs(found.B - expected.B), 0, 3);
            }
        }
    }
}
