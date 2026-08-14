using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void ExtensionlessZipImageReusesTheBudgetedSniffPayload() {
        byte[] image = CreatePngWithC2paManifest(CreateManifestStore());
        byte[] package = CreateCompressedZip(("media/extensionless", image));
        var inspectionOptions = new OfficeProvenanceOptions { MaxExpandedContainerBytes = image.Length };
        var removalOptions = new OfficeProvenanceRemovalOptions();
        removalOptions.Limits.MaxExpandedContainerBytes = image.Length * 2L;

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(package, "package.zip", inspectionOptions);
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "package.zip", removalOptions);

        Assert.True(report.HasC2paManifest);
        Assert.True(result.WasChanged);
        Assert.Empty(OfficeProvenanceInspector.Inspect(ReadZipEntry(result.ToArray(), "media/extensionless"), "image.png").Evidence);
    }

    [Fact]
    public void WebpXmpBeforeTheFinalAnimationFrameIsStructurallyInvalid() {
        byte[] webp = CreateWebp(
            CreateVp8xChunk(advertiseXmp: true),
            CreateRiffChunk("ANMF", new byte[] { 1, 2 }),
            CreateRiffChunk("XMP ", CreateXmpPacket()),
            CreateRiffChunk("ANMF", new byte[] { 3, 4 }));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(webp, "fixture.webp");

        Assert.NotEmpty(result.Before.Evidence);
        Assert.All(result.Before.Evidence, item => Assert.False(item.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(webp, result.ToArray());
    }

    [Fact]
    public void PngCarrierBeforeIendWithTrailingBytesIsStructurallyInvalid() {
        byte[] png = Join(CreatePngWithC2paManifest(CreateManifestStore()), Encoding.ASCII.GetBytes("trailing"));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(png, result.ToArray());
    }

    [Fact]
    public void ZipEntryLookupUsesTheValidatedUnicodePath() {
        byte[] package = CreateZipWithUnicodePathEntry(
            "META-INF/harmless.xml",
            "META-INF/signatures.xml",
            Encoding.UTF8.GetBytes("<signatures/>"));

        Assert.True(OfficeProvenanceZip.HasEntry(
            package,
            path => path.Equals("META-INF/signatures.xml", StringComparison.Ordinal)));
    }

    [Fact]
    public void PackageHelperRewriteHonorsTheAggregateExpandedByteLimit() {
        byte[] package = CreateCompressedZip(
            ("remove.bin", new byte[] { 1 }),
            ("keep.bin", new byte[64]));

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceZip.RemoveEntries(
            package,
            path => path.Equals("remove.bin", StringComparison.Ordinal),
            maximumExpandedBytes: 32));
    }

    [Fact]
    public void TiffC2paRemovalRejectsIfdRewritesThatOverlapRetainedStorage() {
        byte[] manifest = CreateManifestStore();
        const int payloadOffset = 38;
        byte[] tiff = new byte[payloadOffset + manifest.Length];
        tiff[0] = tiff[1] = (byte)'I';
        tiff[2] = 42;
        tiff[4] = 8;
        tiff[8] = 2;
        WriteLittleEndianEntry(tiff, 10, 256, 1, 5, 26);
        WriteLittleEndianEntry(tiff, 22, 0xCD41, 7, manifest.Length, payloadOffset);
        Buffer.BlockCopy(manifest, 0, tiff, payloadOffset, manifest.Length);

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceRemover.Remove(tiff, "fixture.tif"));
    }

    [Fact]
    public void MimetypeValidationUsesThePhysicallyLeadingLocalEntry() {
        byte[] package = CreateStoredZipWithReorderedCentralDirectory(
            ("mimetype", Encoding.ASCII.GetBytes("application/epub+zip")),
            ("content.txt", Encoding.UTF8.GetBytes("content")));

        OfficeProvenanceZip.ValidateMimetypeEntry(package, "application/epub+zip", 10);
    }

    private static byte[] CreateStoredZipWithReorderedCentralDirectory(params (string Name, byte[] Data)[] entries) {
        using var output = new MemoryStream();
        using var writer = new BinaryWriter(output, Encoding.UTF8, leaveOpen: true);
        var records = new List<(byte[] Name, byte[] Data, uint Crc, uint Offset)>();
        foreach ((string name, byte[] data) in entries) {
            byte[] encodedName = Encoding.UTF8.GetBytes(name);
            uint crc = ComputePngCrc(data, 0, data.Length);
            uint offset = checked((uint)output.Position);
            writer.Write(0x04034B50U);
            writer.Write((ushort)20);
            writer.Write((ushort)0x0800);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write(crc);
            writer.Write((uint)data.Length);
            writer.Write((uint)data.Length);
            writer.Write((ushort)encodedName.Length);
            writer.Write((ushort)0);
            writer.Write(encodedName);
            writer.Write(data);
            records.Add((encodedName, data, crc, offset));
        }

        uint centralOffset = checked((uint)output.Position);
        foreach ((byte[] name, byte[] data, uint crc, uint offset) in records.AsEnumerable().Reverse()) {
            writer.Write(0x02014B50U);
            writer.Write((ushort)20);
            writer.Write((ushort)20);
            writer.Write((ushort)0x0800);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write(crc);
            writer.Write((uint)data.Length);
            writer.Write((uint)data.Length);
            writer.Write((ushort)name.Length);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write(0U);
            writer.Write(offset);
            writer.Write(name);
        }
        uint centralSize = checked((uint)output.Position - centralOffset);
        writer.Write(0x06054B50U);
        writer.Write((ushort)0);
        writer.Write((ushort)0);
        writer.Write((ushort)records.Count);
        writer.Write((ushort)records.Count);
        writer.Write(centralSize);
        writer.Write(centralOffset);
        writer.Write((ushort)0);
        writer.Flush();
        return output.ToArray();
    }
}
