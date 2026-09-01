using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfColorRenderingFinalReviewTests {
    [Fact]
    public void ExtractImages_AcceptsDecodeParmsArrayForSingleDctFilterArray() {
        byte[] jpeg = OfficeJpegCodec.Encode(
            OfficeRasterImage.FromRgba32(1, 1, new byte[] { 255, 0, 0, 255 }),
            new OfficeJpegEncodeOptions { Quality = 100, Subsampling = OfficeJpegSubsampling.Y444 });
        byte[] pdf = BuildImagePdf(
            jpeg,
            imageEntries: "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter [/DCTDecode] /DecodeParms [<< /ColorTransform 0 >>]");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(image.IsImageFile);
        Assert.Equal("png", image.FileExtension);
    }

    [Fact]
    public void ImageCollection_DeduplicatesPlacementIntentsOverriddenByImageIntent() {
        var imageDictionary = new PdfDictionary();
        imageDictionary.Items["Type"] = new PdfName("XObject");
        imageDictionary.Items["Subtype"] = new PdfName("Image");
        imageDictionary.Items["Width"] = new PdfNumber(1);
        imageDictionary.Items["Height"] = new PdfNumber(1);
        imageDictionary.Items["ColorSpace"] = new PdfName("DeviceRGB");
        imageDictionary.Items["BitsPerComponent"] = new PdfNumber(8);
        imageDictionary.Items["Intent"] = new PdfName("Perceptual");
        var objects = new Dictionary<int, PdfIndirectObject> {
            [5] = new PdfIndirectObject(5, 0, new PdfStream(imageDictionary, new byte[] { 255, 0, 0 }))
        };
        var xObjects = new PdfDictionary();
        xObjects.Items["Im0"] = new PdfReference(5, 0);
        var resources = new PdfDictionary();
        resources.Items["XObject"] = xObjects;
        PdfImagePlacement[] placements = {
            Placement(OfficeIccRenderingIntent.Saturation),
            Placement(OfficeIccRenderingIntent.AbsoluteColorimetric)
        };

        PdfExtractedImage image = Assert.Single(ResourceResolver.GetImageXObjectsForResources(
            resources,
            objects,
            pageNumber: 1,
            placements));

        Assert.Equal(OfficeIccRenderingIntent.Perceptual, image.RenderingIntent);

        static PdfImagePlacement Placement(OfficeIccRenderingIntent intent) =>
            new PdfImagePlacement(
                1, "Im0", 5, 0,
                1, 0, 0, 1, 0, 0,
                0, 0, 1, 1,
                renderingIntent: intent);
    }

    [Fact]
    public void Move_FailsClosedBeforeRestampingOutputManagedImageBytes() {
        byte[] pdf = BuildImagePdf(
            new byte[] { 255, 0, 0 },
            imageEntries: "/ColorSpace /DeviceRGB /BitsPerComponent 8",
            outputProfile: PdfIccProfiles.SrgbIec6196621);
        PdfDocument document = PdfDocument.Load(pdf);

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            document.Images.Move(Assert.Single(document.Images.Placements()), 10, 0));

        Assert.Contains("output-managed", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Redaction_FailsClosedBeforeNormalizingOutputManagedIccImage() {
        byte[] samples = {
            255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255,
            255, 255, 0, 0, 255, 255, 255, 0, 255, 64, 64, 64
        };
        byte[] pdf = BuildImagePdf(
            samples,
            imageEntries: "/ColorSpace [/ICCBased 7 0 R] /BitsPerComponent 8",
            outputProfile: PdfIccProfiles.SrgbIec6196621,
            sourceProfile: PdfIccProfiles.SrgbIec6196621,
            width: 4,
            height: 2);
        var area = new PdfRedactionArea(1, 20, 30, 20, 20, "icc-image");

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            PdfRedactionApplier.Apply(pdf, new[] { area }));

        Assert.Contains("could not be rewritten safely", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static byte[] BuildImagePdf(
        byte[] imageBytes,
        string imageEntries,
        byte[]? outputProfile = null,
        byte[]? sourceProfile = null,
        int width = 1,
        int height = 1) {
        const string content = "q 40 0 0 20 20 30 cm /Im0 Do Q\n";
        byte[] contentBytes = Encoding.ASCII.GetBytes(content);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(
            output,
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R" +
            (outputProfile == null
                ? string.Empty
                : " /OutputIntents [<< /Type /OutputIntent /S /GTS_PDFA1 /DestOutputProfile 6 0 R >>]") +
            " >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Resources << /XObject << /Im0 5 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "endstream\nendobj\n");
        WriteAscii(
            output,
            "5 0 obj\n<< /Type /XObject /Subtype /Image /Width " + width.ToString(CultureInfo.InvariantCulture) +
            " /Height " + height.ToString(CultureInfo.InvariantCulture) + " " + imageEntries +
            " /Length " + imageBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(imageBytes, 0, imageBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        if (outputProfile != null) WriteProfile(output, 6, outputProfile, includeComponentCount: false);
        if (sourceProfile != null) WriteProfile(output, 7, sourceProfile, includeComponentCount: true);
        WriteAscii(output, "trailer\n<< /Root 1 0 R /Size 8 >>\n%%EOF\n");
        return output.ToArray();
    }

    private static void WriteProfile(Stream output, int objectNumber, byte[] profile, bool includeComponentCount) {
        WriteAscii(
            output,
            objectNumber.ToString(CultureInfo.InvariantCulture) + " 0 obj\n<< " +
            (includeComponentCount ? "/N 3 " : string.Empty) +
            "/Length " + profile.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(profile, 0, profile.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
    }

    private static void WriteAscii(Stream output, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        output.Write(bytes, 0, bytes.Length);
    }
}
