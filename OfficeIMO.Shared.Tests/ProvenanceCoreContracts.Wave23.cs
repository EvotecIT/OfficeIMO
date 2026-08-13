using System.IO.Compression;
using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void ManifestWithoutAssertionStoreIsStructurallyInvalid() {
        byte[] manifest = CreateManifestStore();
        byte[] identity = Encoding.ASCII.GetBytes("c2as");
        int offset = FindSequence(manifest, identity);
        Assert.True(offset >= 0);
        manifest[offset + 3] = (byte)'x';
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", new byte[13]),
            CreatePngChunk("caBX", manifest),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(png, "fixture.png");

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void ExtensionlessSvgSniffTraversesLongLegalPrefixWithinTheAssetBudget() {
        string svg = "<?xml version=\"1.0\"?>" + new string(' ', 17000) + "<!--" + new string('x', 17000) + "-->" +
            "<svg xmlns=\"http://www.w3.org/2000/svg\"><metadata>" + Encoding.UTF8.GetString(CreateXmpPacket()) + "</metadata></svg>";
        byte[] package = CreateCompressedZip(("media/extensionless", Encoding.UTF8.GetBytes(svg)));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(package, "fixture.docx");

        Assert.True(report.HasGenerativeAiDeclaration);
    }

    [Fact]
    public void TiffAcceptsC2paOnlyInThePrimaryIfd() {
        byte[] manifest = CreateManifestStore();
        byte[] tiff = new byte[32 + manifest.Length];
        tiff[0] = tiff[1] = (byte)'I';
        tiff[2] = 42;
        tiff[4] = 8;
        tiff[8] = 1;
        WriteLittleEndianEntry(tiff, 10, 0xCD41, 7, manifest.Length, 32);
        BitConverter.GetBytes(26).CopyTo(tiff, 22);
        BitConverter.GetBytes(0).CopyTo(tiff, 28);
        Buffer.BlockCopy(manifest, 0, tiff, 32, manifest.Length);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(tiff, "fixture.tif");

        Assert.True(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.Empty(result.After.Evidence);
    }

    private static int FindSequence(byte[] data, byte[] sequence) {
        for (int offset = 0; offset <= data.Length - sequence.Length; offset++) {
            bool equal = true;
            for (int index = 0; index < sequence.Length; index++) equal &= data[offset + index] == sequence[index];
            if (equal) return offset;
        }
        return -1;
    }
}
