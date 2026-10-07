using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfAdobeJpegDecodeTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DecodeArrayOwnsAdobeCmykAndYcckPolarity(bool inverted) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", "JpegArithmeticColor");
        byte[] profile = File.ReadAllBytes(Path.Combine(corpus, "profile.icc"));
        string[] files = Directory.GetFiles(corpus, "*.jpg");
        Assert.Equal(32, files.Length);
        foreach (string file in files) {
            byte[] expected = File.ReadAllBytes(file + (inverted ? ".pdf-inverted.srgb" : ".pdf-normal.srgb"));
            byte[] pdf = CreatePdf(File.ReadAllBytes(file), profile, inverted);
            var raster = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(pdf).Pages[0].ToDrawing(), scale: 4D / 3D);
            for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++) {
                int at = (y * 35 + x) * 3;
                var pixel = raster.GetPixel(x * 3 + 1, y * 3 + 1);
                Assert.True(Math.Abs(pixel.R - expected[at]) <= 5 && Math.Abs(pixel.G - expected[at + 1]) <= 5 &&
                    Math.Abs(pixel.B - expected[at + 2]) <= 5 && pixel.A == 255, $"{Path.GetFileName(file)} inverted={inverted} at {x},{y}");
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DeviceCmykImplicitAndExplicitTransformsHaveTheSamePolarity(bool ycck) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", "JpegArithmeticColor");
        byte[] profile = File.ReadAllBytes(Path.Combine(corpus, "profile.icc"));
        foreach (string file in Directory.GetFiles(corpus, ycck ? "b8-c1-*.jpg" : "b8-c0-*.jpg")) {
            byte[] jpeg = File.ReadAllBytes(file);
            foreach (string decode in new[] { "", "/Decode [0 1 0 1 0 1 0 1]", "/Decode [1 0 1 0 1 0 1 0]" }) {
                byte[] implicitPdf = CreatePdf(jpeg, profile, false, "/DeviceCMYK", decode);
                byte[] explicitPdf = CreatePdf(jpeg, profile, false, "/DeviceCMYK", decode +
                    " /DecodeParms << /ColorTransform " + (ycck ? "1" : "0") + " >>");
                var implicitImage = Assert.Single(PdfImageExtractor.ExtractImages(implicitPdf));
                var explicitImage = Assert.Single(PdfImageExtractor.ExtractImages(explicitPdf));
                Assert.Equal("png", implicitImage.FileExtension);
                Assert.Equal("png", explicitImage.FileExtension);
                Assert.Equal(explicitImage.Bytes, implicitImage.Bytes);
                var implicitRaster = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(implicitPdf).Pages[0].ToDrawing());
                var explicitRaster = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(explicitPdf).Pages[0].ToDrawing());
                for (int y = 0; y < implicitRaster.Height; y++) for (int x = 0; x < implicitRaster.Width; x++)
                    Assert.Equal(explicitRaster.GetPixel(x, y), implicitRaster.GetPixel(x, y));
            }
        }
    }

    private static byte[] CreatePdf(byte[] jpeg, byte[] profile, bool inverted, string colorSpace = "[/ICCBased 5 0 R]", string extraEntries = "") {
        using var stream = new MemoryStream();
        void Write(string text) { byte[] bytes = Encoding.ASCII.GetBytes(text); stream.Write(bytes, 0, bytes.Length); }
        var offsets = new List<long> { 0 };
        void Object(int number, string value, byte[]? data = null) {
            offsets.Add(stream.Position);
            Write(number + " 0 obj\n" + value);
            if (data != null) { Write("\nstream\n"); stream.Write(data, 0, data.Length); Write("\nendstream"); }
            Write("\nendobj\n");
        }
        Write("%PDF-1.7\n");
        Object(1, "<< /Type /Catalog /Pages 2 0 R >>");
        Object(2, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>");
        Object(3, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 78.75 42.75] /Resources << /XObject << /Im 4 0 R >> >> /Contents 6 0 R >>");
        Object(4, "<< /Type /XObject /Subtype /Image /Width 35 /Height 19 /BitsPerComponent 8 /ColorSpace " + colorSpace + " " + extraEntries + " /Intent /AbsoluteColorimetric /Filter /DCTDecode " +
            (inverted ? "/Decode [1 0 1 0 1 0 1 0] " : "") + "/Length " + jpeg.Length + " >>", jpeg);
        Object(5, "<< /N 4 /Length " + profile.Length + " >>", profile);
        byte[] content = Encoding.ASCII.GetBytes("q 78.75 0 0 42.75 0 0 cm /Im Do Q");
        Object(6, "<< /Length " + content.Length + " >>", content);
        long xref = stream.Position;
        Write("xref\n0 7\n0000000000 65535 f \n");
        foreach (long offset in offsets.Skip(1)) Write(offset.ToString("D10", CultureInfo.InvariantCulture) + " 00000 n \n");
        Write("trailer\n<< /Size 7 /Root 1 0 R >>\nstartxref\n" + xref.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n");
        return stream.ToArray();
    }
}
