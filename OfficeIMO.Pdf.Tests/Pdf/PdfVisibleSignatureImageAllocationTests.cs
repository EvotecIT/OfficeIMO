using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfVisibleSignatureImageAllocationTests {
    [Fact]
    public void TaggedJpegAppearancePreservesDctBytesAndAllocatesIccProfile() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        byte[] jpeg = OfficeJpegCodec.Encode(new OfficeRasterImage(2, 1, OfficeColor.Red),
            new OfficeJpegEncodeOptions { Metadata = new OfficeJpegMetadata(icc: profile) });

        var (objects, image) = PrepareAppearance(jpeg, PdfFileVersion.Pdf14);

        Assert.Equal("DCTDecode", image.Dictionary.Get<PdfName>("Filter")?.Name);
        Assert.Equal(jpeg, image.Data);
        AssertIccProfile(image, objects, profile);
    }

    [Theory]
    [InlineData(PdfFileVersion.Pdf14, 8)]
    [InlineData(PdfFileVersion.Pdf15, 16)]
    [InlineData(PdfFileVersion.Pdf20, 16)]
    public void TaggedRgbaAppearanceAllocatesProfileAndMaskAtTargetSampleDepth(PdfFileVersion version, int bits) {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        byte[] png = AddIccProfile(PdfPngTestImages.Create16BitRgbaPng(), profile);

        var (objects, image) = PrepareAppearance(png, version);

        AssertIccProfile(image, objects, profile);
        PdfReference maskReference = Assert.IsType<PdfReference>(image.Dictionary.Items["SMask"]);
        PdfStream mask = Assert.IsType<PdfStream>(objects[maskReference.ObjectNumber].Value);
        Assert.Equal(bits, image.Dictionary.Get<PdfNumber>("BitsPerComponent")!.Value);
        Assert.Equal(bits, mask.Dictionary.Get<PdfNumber>("BitsPerComponent")!.Value);
        Assert.Equal(bits, image.Dictionary.Get<PdfDictionary>("DecodeParms")!.Get<PdfNumber>("BitsPerComponent")!.Value);
        Assert.Equal(bits, mask.Dictionary.Get<PdfDictionary>("DecodeParms")!.Get<PdfNumber>("BitsPerComponent")!.Value);
        Assert.Equal(bits == 16 ? new byte[] { 0, 0x12, 0x34, 0x80, 0, 0xff, 0xff } : new byte[] { 0, 18, 128, 255 },
            FlateDecoder.Decode(image.Data));
        Assert.Equal(bits == 16 ? new byte[] { 0, 0x40, 0 } : new byte[] { 0, 64 }, FlateDecoder.Decode(mask.Data));
    }

    [Theory]
    [InlineData("signature")]
    [InlineData("stamp")]
    [InlineData("annotation")]
    public void CatalogVersionUpgradePreservesSixteenBitImageAndAlpha(string path) {
        byte[][] sourceObjects = {
            PdfPageExtractor.WrapObject(1, Encoding.ASCII.GetBytes("<< /Type /Catalog /Pages 2 0 R /Version /1.7 >>\n")),
            PdfPageExtractor.WrapObject(2, Encoding.ASCII.GetBytes("<< /Type /Pages /Count 1 /Kids [3 0 R] >>\n")),
            PdfPageExtractor.WrapObject(3, Encoding.ASCII.GetBytes("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 800] /Resources << >> >>\n"))
        };
        byte[] source = PdfFileAssembler.Assemble(sourceObjects, 1, 0, PdfFileVersion.Pdf14);
        byte[] png = PdfPngTestImages.Create16BitRgbaPng();
        byte[] result = path switch {
            "signature" => PdfIncrementalUpdater.PrepareExternalSignature(source,
                new PdfExternalSignatureOptions {
                    VisibleAppearance = new PdfVisibleSignatureAppearanceOptions { ImageBytes = png, ShowText = false },
                    ReservedSignatureContentsBytes = 512
                }).PreparedPdf,
            "stamp" => PdfDocument.Load(source).Stamp.Image(png).ToBytes(),
            _ => PdfDocument.Load(source).Annotations.AddStamp(new PdfStampAnnotationOptions { ImageBytes = png }).Bytes
        };
        var objects = PdfSyntax.ParseObjects(result).Map;
        PdfStream image = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfStream>(),
            stream => stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Image" && stream.Dictionary.Items.ContainsKey("SMask"));
        PdfReference maskReference = Assert.IsType<PdfReference>(image.Dictionary.Items["SMask"]);
        PdfStream mask = Assert.IsType<PdfStream>(objects[maskReference.ObjectNumber].Value);
        Assert.Equal(16, image.Dictionary.Get<PdfNumber>("BitsPerComponent")!.Value);
        Assert.Equal(16, mask.Dictionary.Get<PdfNumber>("BitsPerComponent")!.Value);
        Assert.Equal(new byte[] { 0, 0x12, 0x34, 0x80, 0, 0xff, 0xff }, FlateDecoder.Decode(image.Data));
        Assert.Equal(new byte[] { 0, 0x40, 0 }, FlateDecoder.Decode(mask.Data));
    }

    private static (Dictionary<int, PdfIndirectObject> Objects, PdfStream Image) PrepareAppearance(byte[] imageBytes, PdfFileVersion version) {
        byte[] source = PdfDocument.Create(new PdfOptions { FileVersion = version })
            .Paragraph(paragraph => paragraph.Text("Signature image allocation source"))
            .ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(source,
            new PdfExternalSignatureOptions {
                FieldName = "ImageApproval",
                VisibleAppearance = new PdfVisibleSignatureAppearanceOptions { ImageBytes = imageBytes, ShowText = false },
                ReservedSignatureContentsBytes = 512
            });
        Assert.True(preparation.PreparedPdf.AsSpan(0, source.Length).SequenceEqual(source));
        var objects = PdfSyntax.ParseObjects(preparation.PreparedPdf).Map;
        PdfStream appearance = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfStream>(),
            stream => stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Form");
        PdfDictionary resources = Assert.IsType<PdfDictionary>(appearance.Dictionary.Items["Resources"]);
        PdfDictionary xObjects = Assert.IsType<PdfDictionary>(resources.Items["XObject"]);
        PdfReference imageReference = Assert.IsType<PdfReference>(xObjects.Items["Im1"]);
        return (objects, Assert.IsType<PdfStream>(objects[imageReference.ObjectNumber].Value));
    }

    private static void AssertIccProfile(PdfStream image, Dictionary<int, PdfIndirectObject> objects, byte[] expected) {
        PdfArray colorSpace = Assert.IsType<PdfArray>(image.Dictionary.Items["ColorSpace"]);
        Assert.Equal("ICCBased", Assert.IsType<PdfName>(colorSpace.Items[0]).Name);
        PdfReference profileReference = Assert.IsType<PdfReference>(colorSpace.Items[1]);
        PdfStream profile = Assert.IsType<PdfStream>(objects[profileReference.ObjectNumber].Value);
        Assert.Equal(3, profile.Dictionary.Get<PdfNumber>("N")!.Value);
        Assert.Equal("DeviceRGB", profile.Dictionary.Get<PdfName>("Alternate")?.Name);
        Assert.Equal(expected, StreamDecoder.Decode(profile.Dictionary, profile.Data, objects));
    }

    private static byte[] AddIccProfile(byte[] png, byte[] profile) {
        byte[] compressed = PdfPngTestImages.BuildStoredZlib(profile);
        byte[] chunk = new byte[compressed.Length + 5];
        Encoding.ASCII.GetBytes("RGB").CopyTo(chunk, 0);
        compressed.CopyTo(chunk, 5); // Profile name, NUL, compression method zero.
        using var output = new MemoryStream();
        const int headerEnd = 8 + 12 + 13; // PNG signature and IHDR chunk.
        output.Write(png, 0, headerEnd);
        PdfPngTestImages.WritePngChunk(output, "iCCP", chunk);
        output.Write(png, headerEnd, png.Length - headerEnd);
        return output.ToArray();
    }
}
