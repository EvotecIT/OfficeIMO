using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDisplayRasterTests {
    [Fact]
    public void InProcessDisplayPixelsMatchTheDecodedDisplayPng() {
        const string content = "0.2 0.4 0.8 rg 10 10 120 80 re f /Half gs 0.9 0.1 0.1 rg 60 40 120 80 re f BT /F1 18 Tf 0 0 0 rg 20 150 Td (Print sheet) Tj ET";
        PdfDocument document = PdfDocument.Load(Encoding.ASCII.GetBytes(string.Join("\n",
            "%PDF-1.4",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R " +
                "/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> /ExtGState << /Half << /ca 0.5 >> >> >> >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF")));
        var options = new PdfPageDisplayOptions { Scale = 150 / 72D };

        (OfficeRasterImage raster, PdfPageRenderResult result) = document.Render.DisplayRaster(1, options, default);
        PdfPageRenderResult png = document.Render.DisplayPage(1, options);

        Assert.True(OfficeRasterImageDecoder.TryDecode(png.Bytes!, out OfficeRasterImage? decoded));
        Assert.Equal((decoded!.Width, decoded.Height), (raster.Width, raster.Height));
        Assert.Equal((png.Width, png.Height), (result.Width, result.Height));
        Assert.Equal(decoded.PixelBuffer, raster.PixelBuffer);
        Assert.Equal(png.Diagnostics, result.Diagnostics);
    }
}
