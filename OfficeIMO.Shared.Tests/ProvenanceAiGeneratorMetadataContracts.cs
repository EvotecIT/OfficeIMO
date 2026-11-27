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
    [InlineData("Midjourney")]
    [InlineData("DALL·E 3")]
    [InlineData("Adobe Firefly (Image 3)")]
    [InlineData("ComfyUI")]
    [InlineData("Image generated by Stable Diffusion XL")]
    public void InspectReportsXmpCreatorToolNamingAnAiGenerator(string creatorTool) {
        byte[] png = CreatePng(CreateXmpChunk($"<rdf:Description rdf:about=\"\" xmp:CreatorTool=\"{creatorTool}\"/>"));

        OfficeProvenanceEvidence evidence = Assert.Single(OfficeProvenanceInspector.Inspect(png, "generated.png").Evidence);

        Assert.Equal(OfficeProvenanceCarrierKind.AiGeneratorMetadata, evidence.Carrier);
        Assert.Equal(creatorTool, evidence.Value);
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

        Assert.Equal(new[] { "Midjourney v6", "UserComment" }, evidence.Select(item => item.Value));
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
        string settings = "Steps: 28, Sampler: DPM++ 2M, CFG scale: 5, Seed: 9";
        foreach (Encoding encoding in new[] { Encoding.BigEndianUnicode, Encoding.Unicode }) {
            byte[] comment = Concat(Encoding.ASCII.GetBytes("UNICODE\0"), encoding.GetBytes(settings));
            byte[] png = CreatePng(CreateChunk("eXIf", CreateExif("Darktable", Latin1.GetString(comment))));

            Assert.Equal("UserComment", Assert.Single(OfficeProvenanceInspector.Inspect(png, "generated.png").Evidence).Value);
        }
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
