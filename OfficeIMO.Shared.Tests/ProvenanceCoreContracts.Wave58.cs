using System;
using System.IO;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void WebpWithNonzeroChunkPaddingIsPreservedByStrictRemoval() {
        byte[] image = CreateRiffChunk("VP8 ", new byte[] { 1 });
        image[image.Length - 1] = 0x7f;
        byte[] webp = CreateWebp(image, CreateRiffChunk("C2PA", CreateManifestStore()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(webp, "fixture.webp");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(webp, result.ToArray());
    }

    [Fact]
    public void ZipCentralDirectorySignatureBlocksDefaultRemoval() {
        byte[] package = AddCentralDirectorySignature(CreateZip(
            ("META-INF/content_credential.c2pa", CreateManifestStore()),
            ("keep.txt", new byte[] { 1, 2, 3 })));

        Assert.Throws<InvalidOperationException>(() => OfficeProvenanceRemover.Remove(package, "fixture.zip"));
    }

    private static byte[] AddCentralDirectorySignature(byte[] package) {
        int endOffset = -1;
        for (int offset = package.Length - 22; offset >= 0; offset--) {
            if (BitConverter.ToUInt32(package, offset) != 0x06054B50U) continue;
            int commentLength = BitConverter.ToUInt16(package, offset + 20);
            if (offset + 22 + commentLength == package.Length) { endOffset = offset; break; }
        }
        Assert.True(endOffset >= 0);
        uint centralSize = BitConverter.ToUInt32(package, endOffset + 12);
        uint centralOffset = BitConverter.ToUInt32(package, endOffset + 16);
        int insertionOffset = checked((int)(centralOffset + centralSize));
        Assert.Equal(endOffset, insertionOffset);

        byte[] signature = { 0x50, 0x4b, 0x05, 0x05, 0x03, 0x00, 0x10, 0x20, 0x30 };
        byte[] result = new byte[package.Length + signature.Length];
        Buffer.BlockCopy(package, 0, result, 0, insertionOffset);
        Buffer.BlockCopy(signature, 0, result, insertionOffset, signature.Length);
        Buffer.BlockCopy(package, insertionOffset, result, insertionOffset + signature.Length, package.Length - insertionOffset);
        BitConverter.GetBytes(centralSize + (uint)signature.Length).CopyTo(result, endOffset + signature.Length + 12);
        return result;
    }
}
