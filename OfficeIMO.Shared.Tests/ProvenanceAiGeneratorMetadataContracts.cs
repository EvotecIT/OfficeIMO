using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class ProvenanceAiGeneratorMetadataContracts {
    private static readonly Encoding Latin1 = Encoding.GetEncoding(28591);
    private static readonly byte[] PngSignature = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

    [Theory]
    [InlineData("tEXt", "parameters")]
    [InlineData("tEXt", "prompt")]
    [InlineData("tEXt", "workflow")]
    [InlineData("zTXt", "parameters")]
    [InlineData("iTXt", "invokeai_metadata")]
    [InlineData("tEXt", "sd-metadata")]
    public void InspectReportsAiGeneratorTextChunkAsDeclaration(string chunkType, string keyword) {
        byte[] png = CreatePng(CreateTextChunk(chunkType, keyword, "a lighthouse at dusk, Steps: 30, Seed: 42"));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(png, "generated.png");

        OfficeProvenanceEvidence evidence = Assert.Single(report.Evidence);
        Assert.Equal(OfficeProvenanceCarrierKind.AiGeneratorMetadata, evidence.Carrier);
        Assert.Equal(keyword, evidence.Value);
        Assert.Equal($"PNG/{chunkType}-{keyword}@{PngSignature.Length + 25}", evidence.Location);
        Assert.True(evidence.IsStructurallyValid);
        Assert.Equal(OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia, evidence.DigitalSourceKind);
        Assert.True(report.HasGenerativeAiDeclaration);
    }

    [Fact]
    public void RemoveDropsOnlyAiGeneratorChunksAndKeepsImageBytes() {
        byte[] title = CreateTextChunk("tEXt", "Title", "Harbour");
        byte[] density = CreateChunk("pHYs", new byte[] { 0, 0, 0x0B, 0x13, 0, 0, 0x0B, 0x13, 1 });
        byte[] prompt = CreateTextChunk("tEXt", "prompt", "{\"3\":{\"class_type\":\"KSampler\"}}");
        byte[] workflow = CreateTextChunk("iTXt", "workflow", "{\"nodes\":[]}");
        byte[] source = CreatePng(title, prompt, density, workflow);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(source, "generated.png", new OfficeProvenanceRemovalOptions());

        Assert.Equal(CreatePng(title, density), result.ToArray());
        Assert.Equal(new[] { "prompt", "workflow" }, result.Changes.Select(change => change.Location.Split('-', '@')[1]));
        Assert.All(result.Changes, change => Assert.Equal(OfficeProvenanceCarrierKind.AiGeneratorMetadata, change.Carrier));
        Assert.Equal(prompt.Length + workflow.Length, result.Changes.Sum(change => change.RemovedBytes));
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void RemoveKeepsAiGeneratorChunksWhenSourceDeclarationsAreKept() {
        byte[] source = CreatePng(CreateTextChunk("tEXt", "parameters", "portrait, Steps: 20"));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(
            source, "generated.png", new OfficeProvenanceRemovalOptions { RemoveAiSourceMetadata = false });

        Assert.False(result.WasChanged);
        Assert.Equal(source, result.ToArray());
    }

    [Theory]
    [InlineData("Parameters")]
    [InlineData("parameters2")]
    [InlineData("Comment")]
    [InlineData("Software")]
    public void InspectIgnoresOtherTextKeywords(string keyword) {
        byte[] png = CreatePng(CreateTextChunk("tEXt", keyword, "Steps: 20"));

        Assert.Empty(OfficeProvenanceInspector.Inspect(png, "image.png").Evidence);
    }

    [Fact]
    public void StrictRemovalPreservesAiGeneratorChunkWithBadCrc() {
        byte[] chunk = CreateTextChunk("tEXt", "parameters", "portrait, Steps: 20");
        chunk[chunk.Length - 1] ^= 0xFF;
        byte[] source = CreatePng(chunk);

        OfficeProvenanceEvidence evidence = Assert.Single(OfficeProvenanceInspector.Inspect(source, "generated.png").Evidence);
        Assert.False(evidence.IsStructurallyValid);

        OfficeProvenanceRemovalResult strict = OfficeProvenanceRemover.Remove(source, "generated.png", new OfficeProvenanceRemovalOptions());
        Assert.Equal(source, strict.ToArray());

        var destructive = new OfficeProvenanceRemovalOptions { RequireStructurallyValidCarrier = false };
        Assert.Equal(CreatePng(), OfficeProvenanceRemover.Remove(source, "generated.png", destructive).ToArray());
    }

    [Fact]
    public void InspectChargesInflatedTextAgainstTheExpandedByteBudget() {
        byte[] text = Concat(Encoding.ASCII.GetBytes("parameters"), new byte[] { 0, 0 }, Zlib(new byte[65536]));
        byte[] png = CreatePng(CreateChunk("zTXt", text));

        Assert.True(Assert.Single(OfficeProvenanceInspector.Inspect(png, "generated.png").Evidence).IsStructurallyValid);

        var limited = new OfficeProvenanceOptions { MaxExpandedContainerBytes = 1024 };
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(png, "generated.png", limited));
    }

    [Fact]
    public void TextAllowanceAdmitsEmptyStreamsAndChargesExceededAndFailedInflation() {
        byte[] empty = CreatePng(CreateChunk("zTXt", Concat(Encoding.ASCII.GetBytes("parameters"), new byte[] { 0, 0 }, EmptyZlib)));
        Assert.True(OfficePngContainerValidator.TryValidate(empty, default, int.MaxValue, 0, out _, out _, out long emptyBytes));
        Assert.Equal(0, emptyBytes);

        const int cap = 1024 * 1024;
        byte[] large = CreatePng(CreateChunk("zTXt", Concat(Encoding.ASCII.GetBytes("parameters"), new byte[] { 0, 0 }, Zlib(new byte[cap + 1]))));
        Assert.False(OfficePngContainerValidator.TryValidate(large, default, int.MaxValue, cap, out _, out _, out long largeBytes));
        Assert.True(largeBytes > cap);

        byte[] stream = Zlib(new byte[4096]);
        stream[stream.Length - 1] ^= 0xFF;
        byte[] corrupt = CreatePng(CreateChunk("zTXt", Concat(Encoding.ASCII.GetBytes("parameters"), new byte[] { 0, 0 }, stream)));
        Assert.False(OfficePngContainerValidator.TryValidate(corrupt, default, int.MaxValue, 1 << 20, out _, out _, out long corruptBytes));
        Assert.True(corruptBytes >= 4096);
    }

    // A final stored block with no data. Built by hand because DeflateStream writes nothing for empty input on some platforms.
    private static readonly byte[] EmptyZlib = { 0x78, 0x9C, 0x03, 0x00, 0x00, 0x00, 0x00, 0x01 };

    private static byte[] Zlib(byte[] data) {
        using var output = new MemoryStream();
        output.WriteByte(0x78);
        output.WriteByte(0x9C);
        using (var deflate = new System.IO.Compression.DeflateStream(output, System.IO.Compression.CompressionLevel.Optimal, leaveOpen: true)) {
            deflate.Write(data, 0, data.Length);
        }
        uint a = 1, b = 0;
        foreach (byte value in data) {
            a = (a + value) % 65521;
            b = (b + a) % 65521;
        }
        uint adler = b << 16 | a;
        output.WriteByte((byte)(adler >> 24));
        output.WriteByte((byte)(adler >> 16));
        output.WriteByte((byte)(adler >> 8));
        output.WriteByte((byte)adler);
        return output.ToArray();
    }

    [Theory]
    [InlineData("Midjourney", "Midjourney")]
    [InlineData("DALL·E 3", "DALL·E")]
    [InlineData("Adobe Firefly (Image 3)", "Adobe Firefly")]
    [InlineData("ComfyUI", "ComfyUI")]
    [InlineData("Image generated by Stable Diffusion XL for a private client", "Stable Diffusion")]
    public void InspectReportsXmpCreatorToolNamingAnAiGenerator(string creatorTool, string product) {
        byte[] png = CreatePng(CreateXmpChunk($"<rdf:Description rdf:about=\"\" xmp:CreatorTool=\"{creatorTool}\"/>"));

        OfficeProvenanceEvidence evidence = Assert.Single(OfficeProvenanceInspector.Inspect(png, "generated.png").Evidence);

        Assert.Equal(OfficeProvenanceCarrierKind.AiGeneratorMetadata, evidence.Carrier);
        Assert.Equal(product, evidence.Value);
        Assert.EndsWith("/xmp:CreatorTool[0]", evidence.Location);
        Assert.True(evidence.IsStructurallyValid);
        Assert.Equal(OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia, evidence.DigitalSourceKind);
    }

    [Theory]
    [InlineData("Adobe Photoshop 25.0 (Macintosh)")]
    [InlineData("Firefly Camera")]
    [InlineData("Gemini Raw Converter")]
    [InlineData("MidjourneyFan Tools")]
    [InlineData("Imagen Studio")]
    public void InspectIgnoresCreatorToolsThatAreNotAiGenerators(string creatorTool) {
        byte[] png = CreatePng(CreateXmpChunk($"<rdf:Description rdf:about=\"\" xmp:CreatorTool=\"{creatorTool}\"/>"));

        Assert.Empty(OfficeProvenanceInspector.Inspect(png, "image.png").Evidence);
    }

    [Fact]
    public void RemoveDropsAiGeneratorXmpFieldsAndKeepsOtherProperties() {
        byte[] png = CreatePng(CreateXmpChunk(
            "<rdf:Description rdf:about=\"\" xmp:CreatorTool=\"Midjourney\" xmp:CreateDate=\"2026-11-01\">" +
            "<tiff:Software>DALL-E</tiff:Software><dc:format>image/png</dc:format></rdf:Description>"));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "generated.png", new OfficeProvenanceRemovalOptions());

        Assert.Equal(new[] { "xmp:CreatorTool", "tiff:Software" },
            result.Changes.Select(change => change.Location.Substring(change.Location.LastIndexOf('/') + 1).Split('[')[0]));
        Assert.Empty(result.After.Evidence);
        string text = Encoding.UTF8.GetString(result.ToArray());
        Assert.DoesNotContain("Midjourney", text);
        Assert.DoesNotContain("DALL-E", text);
        Assert.Contains("2026-11-01", text);
        Assert.Contains("image/png", text);
    }

    [Fact]
    public void RemoveKeepsAiGeneratorXmpFieldsWhenSourceDeclarationsAreKept() {
        byte[] png = CreatePng(CreateXmpChunk("<rdf:Description rdf:about=\"\" xmp:CreatorTool=\"Midjourney\"/>"));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(
            png, "generated.png", new OfficeProvenanceRemovalOptions { RemoveAiSourceMetadata = false });

        Assert.False(result.WasChanged);
    }

    [Fact]
    public void InspectReportsExifSoftwareAndGenerationSettingsInPngExif() {
        byte[] png = CreatePng(CreateChunk("eXIf", CreateExif("Midjourney v6", "ASCII\0\0\0Steps: 20, Sampler: Euler a, CFG scale: 7, Seed: 1")));

        OfficeProvenanceEvidence[] evidence = OfficeProvenanceInspector.Inspect(png, "generated.png").Evidence.ToArray();

        Assert.Equal(new[] { "Midjourney", "UserComment" }, evidence.Select(item => item.Value));
        Assert.Equal(new[] { "IFD0/Software", "ExifIFD/UserComment" }, evidence.Select(item => item.Location.Substring(item.Location.IndexOf('/', 4) + 1)));
        Assert.All(evidence, item => {
            Assert.Equal(OfficeProvenanceCarrierKind.AiGeneratorMetadata, item.Carrier);
            Assert.True(item.IsStructurallyValid);
        });
    }

    [Fact]
    public void RemoveDeletesExifGeneratorEntriesInPlaceAndKeepsOtherTags() {
        byte[] exif = CreateExif("Midjourney v6", "ASCII\0\0\0Steps: 20, Sampler: Euler a, CFG scale: 7, Seed: 1");
        byte[] source = CreatePng(CreateChunk("eXIf", exif));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(source, "generated.png", new OfficeProvenanceRemovalOptions());

        byte[] output = result.ToArray();
        Assert.Equal(source.Length, output.Length);
        Assert.Equal(2, result.Changes.Count);
        Assert.Empty(result.After.Evidence);
        string text = Encoding.ASCII.GetString(output);
        Assert.DoesNotContain("Midjourney", text);
        Assert.DoesNotContain("Steps:", text);
        Assert.Contains("Cam", text);
        // The primary IFD keeps Make and the Exif pointer; the Exif IFD is empty. No other offset moved.
        int tiff = PngSignature.Length + 25 + 8;
        Assert.Equal(2, output[tiff + 8]);
        Assert.Equal(0x010F, output[tiff + 10] | output[tiff + 11] << 8);
        Assert.Equal(0x8769, output[tiff + 22] | output[tiff + 23] << 8);
    }

    [Theory]
    [InlineData("Adobe Photoshop 25.0", "ASCII\0\0\0A walk along the harbour")]
    [InlineData("GIMP 2.10", "UNICODE\0")]
    public void InspectIgnoresOrdinaryExifSoftwareAndComments(string software, string comment) {
        byte[] png = CreatePng(CreateChunk("eXIf", CreateExif(software, comment)));

        Assert.Empty(OfficeProvenanceInspector.Inspect(png, "photo.png").Evidence);
    }

    [Fact]
    public void InspectDecodesUnicodeUserCommentInEitherByteOrder() {
        // Non-Latin text after the product name defeats a zero-byte guess, so the byte-order mark must decide when present.
        string settings = "Midjourney " + new string('\u0100', 20);
        foreach (Encoding encoding in new[] { Encoding.BigEndianUnicode, Encoding.Unicode }) {
            byte[] comment = Concat(Encoding.ASCII.GetBytes("UNICODE\0"), encoding.GetPreamble(), encoding.GetBytes(settings));
            byte[] png = CreatePng(CreateChunk("eXIf", CreateExif("Darktable", Latin1.GetString(comment))));

            Assert.Equal("UserComment", Assert.Single(OfficeProvenanceInspector.Inspect(png, "generated.png").Evidence).Value);
        }
        settings = "Steps: 28, Sampler: DPM++ 2M, CFG scale: 5, Seed: 9";
        foreach (Encoding encoding in new[] { Encoding.BigEndianUnicode, Encoding.Unicode }) {
            byte[] comment = Concat(Encoding.ASCII.GetBytes("UNICODE\0"), encoding.GetBytes(settings));
            byte[] png = CreatePng(CreateChunk("eXIf", CreateExif("Darktable", Latin1.GetString(comment))));

            Assert.Equal("UserComment", Assert.Single(OfficeProvenanceInspector.Inspect(png, "generated.png").Evidence).Value);
        }
    }

    [Fact]
    public void BrokenLinkToALaterIfdStillReportsPrimaryFieldsAsMalformed() {
        byte[] exif = CreateExif("Midjourney", "ASCII\0\0\0holiday");
        // IFD0's next-IFD pointer follows its three entries; point it past the end of the block.
        BitConverter.GetBytes((uint)(exif.Length + 100)).CopyTo(exif, 8 + 2 + 3 * 12);
        byte[] source = CreatePng(CreateChunk("eXIf", exif));

        OfficeProvenanceEvidence evidence = Assert.Single(OfficeProvenanceInspector.Inspect(source, "generated.png").Evidence);
        Assert.Equal("Midjourney", evidence.Value);
        Assert.False(evidence.IsStructurallyValid);

        Assert.False(OfficeProvenanceRemover.Remove(source, "generated.png", new OfficeProvenanceRemovalOptions()).WasChanged);
        OfficeProvenanceRemovalResult destructive = OfficeProvenanceRemover.Remove(
            source, "generated.png", new OfficeProvenanceRemovalOptions { RequireStructurallyValidCarrier = false });
        Assert.Single(destructive.Changes);
        Assert.DoesNotContain("Midjourney", Encoding.ASCII.GetString(destructive.ToArray()));
    }

    [Fact]
    public void BigEndianBigTiffReadsAFourByteExifPointerAtItsDeclaredWidth() {
        byte[] software = Encoding.ASCII.GetBytes("Midjourney\0");
        byte[] comment = Encoding.ASCII.GetBytes("ASCII\0\0\0Steps: 20, Sampler: Euler, Seed: 1");
        const int ifd0 = 16, entry = 20;
        int exifIfd = ifd0 + 8 + 2 * entry + 8;
        int softwareOffset = exifIfd + 8 + entry + 8;
        int commentOffset = softwareOffset + software.Length;
        var tiff = new byte[commentOffset + comment.Length];
        tiff[0] = (byte)'M'; tiff[1] = (byte)'M'; tiff[3] = 43; tiff[5] = 8;
        WriteBigEndianUInt64(tiff, 8, ifd0);
        WriteBigEndianUInt64(tiff, ifd0, 2);
        WriteBigEndianEntry(tiff, ifd0 + 8, 0x0131, 2, (ulong)software.Length);
        WriteBigEndianUInt64(tiff, ifd0 + 8 + 12, (ulong)softwareOffset);
        WriteBigEndianEntry(tiff, ifd0 + 8 + entry, 0x8769, 4, 1);
        WriteBigEndianUInt32(tiff, ifd0 + 8 + entry + 12, (uint)exifIfd); // four-byte LONG, then four bytes of padding
        WriteBigEndianUInt64(tiff, exifIfd, 1);
        WriteBigEndianEntry(tiff, exifIfd + 8, 0x9286, 7, (ulong)comment.Length);
        WriteBigEndianUInt64(tiff, exifIfd + 8 + 12, (ulong)commentOffset);
        software.CopyTo(tiff, softwareOffset);
        comment.CopyTo(tiff, commentOffset);

        OfficeProvenanceEvidence[] evidence = OfficeProvenanceInspector.Inspect(CreatePng(CreateChunk("eXIf", tiff)), "generated.png").Evidence.ToArray();

        Assert.Equal(new[] { "Midjourney", "UserComment" }, evidence.Select(item => item.Value));
    }

    [Fact]
    public void SoftwareCountPastTheEndIsIgnoredInsteadOfFailingInspection() {
        byte[] exif = CreateExif("Midjourney", "ASCII\0\0\0holiday");
        BitConverter.GetBytes(10_000u).CopyTo(exif, 8 + 2 + 12 + 4);
        byte[] png = CreatePng(CreateChunk("eXIf", exif));

        Assert.Empty(OfficeProvenanceInspector.Inspect(png, "photo.png").Evidence);
    }

    [Fact]
    public void DecodedGeneratorValuesDrawFromTheExpandedByteBudget() {
        byte[] png = CreatePng(CreateChunk("eXIf", CreateExif("Midjourney v6 with a longer description", "ASCII\0\0\0holiday")));

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(png, "generated.png",
            new OfficeProvenanceOptions { MaxExpandedContainerBytes = 16 }));
    }

    [Fact]
    public void GeneratorPropertyWithItsOwnAttributesIsNotTreatedAsADeclarationOnlyPacket() {
        static byte[] Packet(string property) => Encoding.UTF8.GetBytes(
            "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">" +
            "<rdf:Description rdf:about=\"\" xmlns:xmp=\"http://ns.adobe.com/xap/1.0/\" xmlns:private=\"urn:private\">" + property +
            "</rdf:Description></rdf:RDF></x:xmpmeta>");

        Assert.True(OfficeProvenanceXmp.HasOnlyDigitalSourceTypes(Packet("<xmp:CreatorTool>Midjourney</xmp:CreatorTool>"), new OfficeProvenanceOptions()));
        Assert.False(OfficeProvenanceXmp.HasOnlyDigitalSourceTypes(
            Packet("<xmp:CreatorTool private:owner=\"Alice\">Midjourney</xmp:CreatorTool>"), new OfficeProvenanceOptions()));
        Assert.False(OfficeProvenanceXmp.HasOnlyDigitalSourceTypes(
            Packet("<xmp:CreatorTool>ComfyUI /home/alice/clients/secret</xmp:CreatorTool>"), new OfficeProvenanceOptions()));
    }

    private static void WriteBigEndianEntry(byte[] data, int offset, ushort tag, ushort type, ulong count) {
        data[offset] = (byte)(tag >> 8); data[offset + 1] = (byte)tag;
        data[offset + 2] = (byte)(type >> 8); data[offset + 3] = (byte)type;
        WriteBigEndianUInt64(data, offset + 4, count);
    }

    private static void WriteBigEndianUInt64(byte[] data, int offset, ulong value) {
        for (int index = 7; index >= 0; index--) { data[offset + index] = (byte)value; value >>= 8; }
    }

    private static void WriteBigEndianUInt32(byte[] data, int offset, uint value) {
        for (int index = 3; index >= 0; index--) { data[offset + index] = (byte)value; value >>= 8; }
    }

    [Fact]
    public void UnmarkedUnicodeCommentIsRecognizedInEitherByteOrderWhateverItsZeroBytes() {
        string text = "Midjourney " + new string('\u0100', 20);
        foreach (Encoding encoding in new[] { Encoding.BigEndianUnicode, Encoding.Unicode }) {
            byte[] comment = Concat(Encoding.ASCII.GetBytes("UNICODE\0"), encoding.GetBytes(text));
            byte[] png = CreatePng(CreateChunk("eXIf", CreateExif("Darktable", Latin1.GetString(comment))));

            Assert.Equal("UserComment", Assert.Single(OfficeProvenanceInspector.Inspect(png, "generated.png").Evidence).Value);
        }
    }

    [Fact]
    public void MalformedSubIfdLinkStillReportsPrimaryFieldsAsMalformed() {
        byte[] exif = CreateExif("Midjourney", "ASCII\0\0\0holiday");
        // Turn Make into a SubIFDs entry (tag 330) with a SHORT pointer type, which is malformed but not a limit.
        BitConverter.GetBytes((ushort)330).CopyTo(exif, 10);
        BitConverter.GetBytes((ushort)3).CopyTo(exif, 12);
        BitConverter.GetBytes(1u).CopyTo(exif, 14);
        byte[] png = CreatePng(CreateChunk("eXIf", exif));

        OfficeProvenanceEvidence evidence = Assert.Single(OfficeProvenanceInspector.Inspect(png, "generated.png").Evidence);
        Assert.Equal("Midjourney", evidence.Value);
        Assert.False(evidence.IsStructurallyValid);
    }

    [Fact]
    public void SoftwareTextSharedWithAGpsEntryIsNotCleared() {
        byte[] exif = CreateExif("Midjourney", "ASCII\0\0\0holiday");
        int software = BitConverter.ToInt32(exif, 22 + 8);
        // Point an extra GPS IFD's GPSMapDatum entry at the Software text by turning the Exif pointer into a GPS pointer and reusing
        // its IFD: tag 0x0012 (GPSMapDatum), ASCII, same count and offset as Software.
        BitConverter.GetBytes((ushort)0x8825).CopyTo(exif, 34);
        int gps = BitConverter.ToInt32(exif, 34 + 8);
        BitConverter.GetBytes((ushort)0x0012).CopyTo(exif, gps + 2);
        BitConverter.GetBytes((ushort)2).CopyTo(exif, gps + 4);
        BitConverter.GetBytes(BitConverter.ToUInt32(exif, 22 + 4)).CopyTo(exif, gps + 6);
        BitConverter.GetBytes((uint)software).CopyTo(exif, gps + 10);
        byte[] source = CreatePng(CreateChunk("eXIf", exif));

        Assert.False(OfficeProvenanceRemover.Remove(source, "generated.png", new OfficeProvenanceRemovalOptions()).WasChanged);
    }

    [Fact]
    public void OrdinaryExifIsNotACarrier() {
        byte[] png = CreatePng(CreateChunk("eXIf", CreateExif("Darktable 4.6", "ASCII\0\0\0holiday")));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(png, "photo.png");
        Assert.Empty(report.Evidence);
        // Only the EXIF text is charged: Software and UserComment (14 + 15 bytes), read by the carrier check and again by inspection.
        // The image itself is not decoded.
        Assert.Equal(2 * (14 + 15), report.ExpandedInspectionBytes);
    }

    [Fact]
    public void StrictRemovalKeepsSoftwareWhoseTextIsSharedWithAnotherEntry() {
        byte[] exif = CreateExif("Midjourney", "ASCII\0\0\0holiday");
        // Point a second IFD0 entry, ImageDescription, at the same out-of-line text by turning Make into tag 270.
        BitConverter.GetBytes((ushort)0x010E).CopyTo(exif, 10);
        BitConverter.GetBytes(BitConverter.ToUInt32(exif, 22 + 4)).CopyTo(exif, 10 + 4);
        BitConverter.GetBytes(BitConverter.ToUInt32(exif, 22 + 8)).CopyTo(exif, 10 + 8);
        byte[] source = CreatePng(CreateChunk("eXIf", exif));

        OfficeProvenanceRemovalResult strict = OfficeProvenanceRemover.Remove(source, "generated.png", new OfficeProvenanceRemovalOptions());
        Assert.False(strict.WasChanged);

        OfficeProvenanceRemovalResult destructive = OfficeProvenanceRemover.Remove(
            source, "generated.png", new OfficeProvenanceRemovalOptions { RequireStructurallyValidCarrier = false });
        Assert.Single(destructive.Changes);
        Assert.Contains("Midjourney", Encoding.ASCII.GetString(destructive.ToArray()));
    }

    /// <summary>
    /// Builds a little-endian EXIF block: IFD0 with Make "Cam" (inline), Software (out of line) and the Exif IFD pointer, and an
    /// Exif IFD with one UserComment whose raw bytes are <paramref name="userComment"/> as Latin-1.
    /// </summary>
    private static byte[] CreateExif(string software, string userComment) {
        byte[] softwareBytes = Encoding.ASCII.GetBytes(software + "\0");
        byte[] commentBytes = Latin1.GetBytes(userComment);
        const int ifd0 = 8;
        const int ifd0Size = 2 + 3 * 12 + 4;
        int exifIfd = ifd0 + ifd0Size;
        const int exifIfdSize = 2 + 12 + 4;
        int softwareOffset = exifIfd + exifIfdSize;
        int commentOffset = softwareOffset + softwareBytes.Length;
        byte[] tiff = new byte[commentOffset + commentBytes.Length];
        tiff[0] = (byte)'I'; tiff[1] = (byte)'I'; tiff[2] = 42; tiff[4] = ifd0;
        tiff[ifd0] = 3;
        WriteEntry(tiff, ifd0 + 2, 0x010F, 2, 4, BitConverter.ToUInt32(Encoding.ASCII.GetBytes("Cam\0"), 0));
        WriteEntry(tiff, ifd0 + 14, 0x0131, 2, (uint)softwareBytes.Length, (uint)softwareOffset);
        WriteEntry(tiff, ifd0 + 26, 0x8769, 4, 1, (uint)exifIfd);
        tiff[exifIfd] = 1;
        WriteEntry(tiff, exifIfd + 2, 0x9286, 7, (uint)commentBytes.Length, (uint)commentOffset);
        softwareBytes.CopyTo(tiff, softwareOffset);
        commentBytes.CopyTo(tiff, commentOffset);
        return tiff;
    }

    private static void WriteEntry(byte[] data, int offset, ushort tag, ushort type, uint count, uint value) {
        BitConverter.GetBytes(tag).CopyTo(data, offset);
        BitConverter.GetBytes(type).CopyTo(data, offset + 2);
        BitConverter.GetBytes(count).CopyTo(data, offset + 4);
        BitConverter.GetBytes(value).CopyTo(data, offset + 8);
    }

    private static byte[] CreateXmpChunk(string description) {
        string xmp = "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
            "xmlns:xmp=\"http://ns.adobe.com/xap/1.0/\" xmlns:tiff=\"http://ns.adobe.com/tiff/1.0/\" " +
            "xmlns:dc=\"http://purl.org/dc/elements/1.1/\">" + description + "</rdf:RDF></x:xmpmeta>";
        byte[] payload = Concat(Encoding.ASCII.GetBytes("XML:com.adobe.xmp"), new byte[] { 0, 0, 0, 0, 0 }, Encoding.UTF8.GetBytes(xmp));
        return CreateChunk("iTXt", payload);
    }

    internal static byte[] CreatePng(params byte[][] ancillary) {
        var parts = new List<byte[]> {
            PngSignature,
            CreateChunk("IHDR", new byte[] { 0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0 })
        };
        parts.AddRange(ancillary);
        parts.Add(CreateChunk("IDAT", new byte[] { 0x78, 0x9C, 0x63, 0x60, 0x60, 0x60, 0x00, 0x00, 0x00, 0x04, 0x00, 0x01 }));
        parts.Add(CreateChunk("IEND", Array.Empty<byte>()));
        return parts.SelectMany(part => part).ToArray();
    }

    /// <summary>Builds a tEXt, zTXt or iTXt chunk with the given keyword.</summary>
    internal static byte[] CreateTextChunk(string type, string keyword, string text) {
        byte[] name = Encoding.ASCII.GetBytes(keyword);
        byte[] value = Encoding.UTF8.GetBytes(text);
        byte[] payload = type switch {
            "tEXt" => Concat(name, new byte[] { 0 }, value),
            // zTXt: keyword, separator, compression method 0, then a zlib stream; detection never reads the stream.
            "zTXt" => Concat(name, new byte[] { 0, 0 }, new byte[] { 0x78, 0x9C, 0x03, 0x00, 0x00, 0x00, 0x00, 0x01 }),
            // iTXt: keyword, separator, uncompressed flag and method, empty language tag and translated keyword, then UTF-8 text.
            "iTXt" => Concat(name, new byte[] { 0, 0, 0, 0, 0 }, value),
            _ => throw new ArgumentOutOfRangeException(nameof(type))
        };
        return CreateChunk(type, payload);
    }

    private static byte[] CreateChunk(string type, byte[] payload) {
        byte[] chunk = new byte[payload.Length + 12];
        WriteBigEndian(chunk, 0, (uint)payload.Length);
        Encoding.ASCII.GetBytes(type).CopyTo(chunk, 4);
        payload.CopyTo(chunk, 8);
        WriteBigEndian(chunk, chunk.Length - 4, Crc(chunk, 4, payload.Length + 4));
        return chunk;
    }

    private static byte[] Concat(params byte[][] parts) => parts.SelectMany(part => part).ToArray();

    private static void WriteBigEndian(byte[] data, int offset, uint value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }

    private static uint Crc(byte[] data, int offset, int count) {
        uint crc = 0xFFFFFFFF;
        for (int index = offset; index < offset + count; index++) {
            crc ^= data[index];
            for (int bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? 0xEDB88320U ^ (crc >> 1) : crc >> 1;
        }
        return crc ^ 0xFFFFFFFF;
    }
}
