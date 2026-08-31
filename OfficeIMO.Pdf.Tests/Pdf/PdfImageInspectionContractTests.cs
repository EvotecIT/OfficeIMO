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

    [Fact]
    public void ImageInspection_ImageDictionaryIntentOverridesPlacementGraphicsState() {
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>",
            StreamObject("q /Saturation ri 10 0 0 10 10 10 cm /Im1 Do Q q /AbsoluteColorimetric ri 10 0 0 10 30 10 cm BI /W 1 /H 1 /BPC 8 /CS /RGB /Intent /Perceptual ID abc EI Q"),
            StreamObject("abc", "/Type /XObject /Subtype /Image /Width 1 /Height 1 /BitsPerComponent 8 /ColorSpace /DeviceRGB /Intent /Perceptual"));
        PdfDocument source = PdfDocument.Open(pdf);

        PdfImagePlacement[] placements = source.Read.ImagePlacements().ToArray();
        PdfExtractedImage[] images = source.Read.Images().ToArray();

        Assert.Equal(2, placements.Length);
        Assert.All(placements, placement => {
            Assert.True(placement.HasAuthoredRenderingIntent);
            Assert.Equal(OfficeIccRenderingIntent.Perceptual, placement.RenderingIntent);
        });
        Assert.Equal(2, images.Length);
        Assert.All(images, image => {
            Assert.True(image.HasAuthoredRenderingIntent);
            Assert.Equal(OfficeIccRenderingIntent.Perceptual, image.RenderingIntent);
        });
    }

    [Fact]
    public void ImageInspection_DistinguishesDefaultAndExplicitNormalBlendModes() {
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /Im1 5 0 R >> /ExtGState << /GS1 6 0 R /GS2 7 0 R >> >> /Contents 4 0 R >>",
            StreamObject("q 10 0 0 10 10 10 cm /Im1 Do Q q /GS1 gs 10 0 0 10 30 10 cm /Im1 Do Q q 10 0 0 10 50 10 cm /Im1 Do Q q /GS2 gs 10 0 0 10 70 10 cm /Im1 Do Q"),
            StreamObject("abc", "/Type /XObject /Subtype /Image /Width 1 /Height 1 /BitsPerComponent 8 /ColorSpace /DeviceRGB"),
            "<< /Type /ExtGState /BM /Normal >>",
            "<< /Type /ExtGState /BM /NotSupported >>");

        PdfImagePlacement[] placements = PdfDocument.Open(pdf).Read.ImagePlacements().ToArray();

        Assert.Equal(4, placements.Length);
        Assert.Null(placements[0].BlendMode);
        Assert.Null(placements[0].AuthoredBlendMode);
        Assert.Equal(OfficeBlendMode.Normal, placements[0].EffectiveBlendMode);
        Assert.Equal(OfficeBlendMode.Normal, placements[1].BlendMode);
        Assert.Equal(OfficeBlendMode.Normal, placements[1].AuthoredBlendMode);
        Assert.Equal(OfficeBlendMode.Normal, placements[1].EffectiveBlendMode);
        Assert.Null(placements[2].BlendMode);
        Assert.Null(placements[2].AuthoredBlendMode);
        Assert.Equal(OfficeBlendMode.Normal, placements[2].EffectiveBlendMode);
        Assert.Null(placements[3].BlendMode);
        Assert.Null(placements[3].AuthoredBlendMode);
        Assert.Equal(OfficeBlendMode.Normal, placements[3].EffectiveBlendMode);
        Assert.True(placements[3].HasUnsupportedBlendMode);
    }

    [Fact]
    public void ImageInspection_PreservesAuthoredBlendModeThroughNestedForms() {
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /Fm1 5 0 R >> /ExtGState << /GS1 7 0 R >> >> /Contents 4 0 R >>",
            StreamObject("q /GS1 gs /Fm1 Do Q"),
            StreamObject("q 10 0 0 10 10 10 cm /Im1 Do Q", "/Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources << /XObject << /Im1 6 0 R >> >>"),
            StreamObject("abc", "/Type /XObject /Subtype /Image /Width 1 /Height 1 /BitsPerComponent 8 /ColorSpace /DeviceRGB"),
            "<< /Type /ExtGState /BM /Screen >>");

        PdfImagePlacement placement = Assert.Single(PdfDocument.Open(pdf).Read.ImagePlacements());

        Assert.Equal(OfficeBlendMode.Screen, placement.BlendMode);
        Assert.Equal(OfficeBlendMode.Screen, placement.AuthoredBlendMode);
        Assert.Equal(OfficeBlendMode.Screen, placement.EffectiveBlendMode);
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
