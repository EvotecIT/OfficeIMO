using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfImageInspectionContractTests {
    [Fact]
    public void ImageInspection_ExposesLowLevelResourceAndPlacementSemantics() {
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /Im1 5 0 R >> /ExtGState << /GS1 6 0 R >> >> /Contents 4 0 R >>",
            StreamObject("q 20 30 100 80 re W n /Perceptual ri /GS1 gs 100 0 0 80 20 30 cm /Im1 Do Q"),
            StreamObject("abc", "/Type /XObject /Subtype /Image /Width 1 /Height 1 /BitsPerComponent 8 /ColorSpace /DeviceRGB /Decode [1 0 1 0 1 0] /DecodeParms << /Predictor 1 >> /Intent /Perceptual /Interpolate true"),
            "<< /Type /ExtGState /ca 0.5 /BM /Multiply >>");
        PdfDocument source = PdfDocument.Open(pdf);

        PdfExtractedImage image = Assert.Single(source.Read.Images());
        PdfImagePlacement placement = Assert.Single(source.Read.ImagePlacements());

        Assert.Equal(OfficeIccRenderingIntent.Perceptual, image.RenderingIntent);
        Assert.True(image.HasAuthoredRenderingIntent);
        Assert.True(image.HasExplicitDecode);
        Assert.True(image.HasDecodeParameters);
        Assert.True(image.Interpolate);
        Assert.Equal(0.5D, placement.ImageOpacity);
        Assert.Equal(0.5D, placement.AuthoredOpacity);
        Assert.Equal(0.5D, placement.Opacity);
        Assert.Equal(OfficeBlendMode.Multiply, placement.BlendMode);
        Assert.Equal(OfficeBlendMode.Multiply, placement.AuthoredBlendMode);
        Assert.Equal(OfficeBlendMode.Multiply, placement.EffectiveBlendMode);
        Assert.False(placement.HasUnsupportedBlendMode);
        Assert.False(placement.HasSoftMask);
        Assert.True(placement.HasAuthoredRenderingIntent);
        Assert.Equal(OfficeIccRenderingIntent.Perceptual, placement.RenderingIntent);
        Assert.NotNull(placement.Clip);
        Assert.True(placement.Clip!.IsRectangle);
        Assert.True(placement.Clip.IsExact);
        Assert.Equal(20D, placement.Clip.X);
        Assert.Equal(190D, placement.Clip.Y);
        Assert.Equal(100D, placement.Clip.Width);
        Assert.Equal(80D, placement.Clip.Height);
        Assert.Empty(placement.Clip.Commands);
        Assert.True(placement.PaintOrder >= 0D);
    }

    [Fact]
    public void ImageInspection_PreservesAuthoredGraphicsStateIntentWhenEffectiveIntentIsDeduplicated() {
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>",
            StreamObject("q 10 0 0 10 10 10 cm /Im1 Do Q q /RelativeColorimetric ri 10 0 0 10 30 10 cm /Im1 Do Q"),
            StreamObject("abc", "/Type /XObject /Subtype /Image /Width 1 /Height 1 /BitsPerComponent 8 /ColorSpace /DeviceRGB"));

        PdfExtractedImage image = Assert.Single(PdfDocument.Open(pdf).Read.Images());

        Assert.Equal(OfficeIccRenderingIntent.RelativeColorimetric, image.RenderingIntent);
        Assert.True(image.HasAuthoredRenderingIntent);
    }

    private static string StreamObject(string content, string additionalDictionary = "") {
        int length = Encoding.ASCII.GetByteCount(content);
        string suffix = string.IsNullOrWhiteSpace(additionalDictionary) ? string.Empty : " " + additionalDictionary;
        return "<< /Length " + length.ToString(CultureInfo.InvariantCulture) + suffix + " >>\nstream\n" + content + "\nendstream";
    }

    private static byte[] BuildPdf(params string[] objects) {
        var builder = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int>(objects.Length);
        for (int i = 0; i < objects.Length; i++) {
            offsets.Add(Encoding.ASCII.GetByteCount(builder.ToString()));
            builder.Append(i + 1).Append(" 0 obj\n").Append(objects[i]).Append("\nendobj\n");
        }

        int xrefOffset = Encoding.ASCII.GetByteCount(builder.ToString());
        builder.Append("xref\n0 ").Append(objects.Length + 1).Append("\n0000000000 65535 f \n");
        for (int i = 0; i < offsets.Count; i++) {
            builder.Append(offsets[i].ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        }
        builder.Append("trailer\n<< /Root 1 0 R /Size ").Append(objects.Length + 1).Append(" >>\nstartxref\n")
            .Append(xrefOffset.ToString(CultureInfo.InvariantCulture)).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(builder.ToString());
    }
}
