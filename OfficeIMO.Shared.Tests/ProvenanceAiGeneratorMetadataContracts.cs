using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class ProvenanceAiGeneratorMetadataContracts {
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
