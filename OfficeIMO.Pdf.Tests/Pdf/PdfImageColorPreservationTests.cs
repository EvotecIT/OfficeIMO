using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfImageColorPreservationTests {
    [Fact]
    public void TaggedJpegPreservesDctAndCompleteApp2Profile() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        byte[] jpeg = OfficeJpegCodec.Encode(new OfficeRasterImage(2, 1, OfficeColor.Red),
            new OfficeJpegEncodeOptions { Metadata = new OfficeJpegMetadata(icc: profile) });
        byte[] pdf = PdfDocument.Create().Image(jpeg, 20, 10).ToBytes();
        var objects = PdfSyntax.ParseObjects(pdf).Map;
        PdfStream image = Assert.Single(Images(objects));
        Assert.Equal("DCTDecode", image.Dictionary.Get<PdfName>("Filter")?.Name);
        Assert.Equal(jpeg, image.Data);
        AssertProfile(image.Dictionary.Items["ColorSpace"], objects, profile);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TaggedPngRetainsProfileAndAlphaWithoutExpandingPalette(bool indexed) {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        byte[] png = CreateTaggedPng(profile, indexed);
        byte[] pdf = PdfDocument.Create().Image(png, 20, 10).ToBytes();
        AssertTaggedPng(pdf, profile, indexed);

        // Stamps use typed dictionaries, unlike the direct writer's string dictionaries.
        byte[] stamped = PdfStamper.StampImage(PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Stamp target")).ToBytes(), png,
            new PdfImageStampOptions { X = 20, Y = 20, Width = 20, Height = 10 });
        AssertTaggedPng(stamped, profile, indexed);
    }

    [Theory]
    [InlineData(PdfFileVersion.Pdf14, 8)]
    [InlineData(PdfFileVersion.Pdf15, 16)]
    [InlineData(PdfFileVersion.Pdf20, 16)]
    public void SixteenBitColorAndAlphaRespectTargetVersion(PdfFileVersion version, int bits) {
        // Reuse the same source across versions to protect the immutable preparation cache.
        byte[] png = PdfPngTestImages.Create16BitRgbaPng();
        byte[] pdf = PdfDocument.Create(new PdfOptions { FileVersion = version }).Image(png, 10, 10).ToBytes();
        var objects = PdfSyntax.ParseObjects(pdf).Map;
        PdfStream image = Images(objects).Single(stream => stream.Dictionary.Items.ContainsKey("SMask"));
        PdfReference maskRef = Assert.IsType<PdfReference>(image.Dictionary.Items["SMask"]);
        PdfStream mask = Assert.IsType<PdfStream>(objects[maskRef.ObjectNumber].Value);
        Assert.Equal(bits, image.Dictionary.Get<PdfNumber>("BitsPerComponent")!.Value);
        Assert.Equal(bits, mask.Dictionary.Get<PdfNumber>("BitsPerComponent")!.Value);
        Assert.Equal(bits, image.Dictionary.Get<PdfDictionary>("DecodeParms")!.Get<PdfNumber>("BitsPerComponent")!.Value);
        Assert.Equal(bits == 16 ? new byte[] { 0, 0x12, 0x34, 0x80, 0, 0xff, 0xff } : new byte[] { 0, 18, 128, 255 },
            FlateDecoder.Decode(image.Data));
        Assert.Equal(bits == 16 ? new byte[] { 0, 0x40, 0 } : new byte[] { 0, 64 }, FlateDecoder.Decode(mask.Data));
    }

    [Fact]
    public void PngRejectsIncompatibleProfileInsteadOfMislabelingSamples() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        Encoding.ASCII.GetBytes("GRAY").CopyTo(profile, 16);
        Assert.False(PdfDocument.TryValidateImageBytes(CreateTaggedPng(profile, false), out _, out string? reason));
        Assert.Contains("compatible", reason);
    }

    private static void AssertTaggedPng(byte[] pdf, byte[] profile, bool indexed) {
        var objects = PdfSyntax.ParseObjects(pdf).Map;
        PdfStream image = Images(objects).Single(stream => stream.Dictionary.Items.ContainsKey("SMask"));
        PdfObject colorSpace = image.Dictionary.Items["ColorSpace"];
        if (indexed) {
            PdfArray palette = Assert.IsType<PdfArray>(colorSpace);
            Assert.Equal("Indexed", Assert.IsType<PdfName>(palette.Items[0]).Name);
            Assert.Equal(1, Assert.IsType<PdfNumber>(palette.Items[2]).Value);
            Assert.Equal(new byte[] { 255, 0, 0, 0, 0, 255 }, Assert.IsType<PdfStringObj>(palette.Items[3]).RawBytes);
            Assert.Equal(2, image.Dictionary.Get<PdfNumber>("BitsPerComponent")!.Value);
            Assert.Equal(1, image.Dictionary.Get<PdfDictionary>("DecodeParms")!.Get<PdfNumber>("Colors")!.Value);
            Assert.Equal(new byte[] { 0, 0x10 }, FlateDecoder.Decode(image.Data));
            colorSpace = palette.Items[1];
        }
        AssertProfile(colorSpace, objects, profile);
        PdfReference maskRef = Assert.IsType<PdfReference>(image.Dictionary.Items["SMask"]);
        PdfStream mask = Assert.IsType<PdfStream>(objects[maskRef.ObjectNumber].Value);
        Assert.Equal(indexed ? new byte[] { 0, 128, 255 } : new byte[] { 0, 128 }, FlateDecoder.Decode(mask.Data));
    }

    private static void AssertProfile(PdfObject colorSpace, Dictionary<int, PdfIndirectObject> objects, byte[] expected) {
        PdfArray icc = Assert.IsType<PdfArray>(colorSpace);
        Assert.Equal("ICCBased", Assert.IsType<PdfName>(icc.Items[0]).Name);
        PdfReference profileRef = Assert.IsType<PdfReference>(icc.Items[1]);
        PdfStream profile = Assert.IsType<PdfStream>(objects[profileRef.ObjectNumber].Value);
        Assert.Equal(3, profile.Dictionary.Get<PdfNumber>("N")!.Value);
        Assert.Equal(expected, StreamDecoder.Decode(profile.Dictionary, profile.Data, objects));
    }

    private static IEnumerable<PdfStream> Images(Dictionary<int, PdfIndirectObject> objects) =>
        objects.Values.Select(item => item.Value).OfType<PdfStream>()
            .Where(stream => stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Image");

    private static byte[] CreateTaggedPng(byte[] profile, bool indexed) {
        using var output = new MemoryStream();
        byte[] signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
        output.Write(signature, 0, signature.Length);
        PdfPngTestImages.WritePngChunk(output, "IHDR", new byte[] {
            0, 0, 0, (byte)(indexed ? 2 : 1), 0, 0, 0, 1, (byte)(indexed ? 2 : 8), (byte)(indexed ? 3 : 6), 0, 0, 0
        });
        byte[] compressed = PdfPngTestImages.BuildStoredZlib(profile);
        byte[] iccp = new byte[compressed.Length + 5];
        Encoding.ASCII.GetBytes("RGB").CopyTo(iccp, 0);
        compressed.CopyTo(iccp, 5); // profile name, NUL, compression method zero
        PdfPngTestImages.WritePngChunk(output, "iCCP", iccp);
        if (indexed) {
            PdfPngTestImages.WritePngChunk(output, "PLTE", new byte[] { 255, 0, 0, 0, 0, 255 });
            PdfPngTestImages.WritePngChunk(output, "tRNS", new byte[] { 128, 255 });
        }
        PdfPngTestImages.WritePngChunk(output, "IDAT", PdfPngTestImages.BuildStoredZlib(
            indexed ? new byte[] { 0, 0x10 } : new byte[] { 0, 12, 34, 56, 128 }));
        PdfPngTestImages.WritePngChunk(output, "IEND", Array.Empty<byte>());
        return output.ToArray();
    }
}
