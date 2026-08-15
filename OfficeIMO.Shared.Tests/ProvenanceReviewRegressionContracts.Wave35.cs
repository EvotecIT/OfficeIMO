using System.IO.Compression;
using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceReviewRegressionContracts {
    [Fact]
    public void ZipRewriteHonorsValidInfoZipUnicodeComments() {
        byte[] rawComment = { 0x82 };
        byte[] unicodeComment = Encoding.UTF8.GetBytes("intended Ω");
        byte[] unicodeExtra = new byte[9 + unicodeComment.Length];
        WriteLittleEndian16(unicodeExtra, 0, 0x6375);
        WriteLittleEndian16(unicodeExtra, 2, checked((ushort)(5 + unicodeComment.Length)));
        unicodeExtra[4] = 1;
        WriteLittleEndian(unicodeExtra, 5, Crc32(rawComment));
        Buffer.BlockCopy(unicodeComment, 0, unicodeExtra, 9, unicodeComment.Length);

        byte[] package;
        using (var stream = new MemoryStream()) {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
                using (Stream manifest = archive.CreateEntry("META-INF/content_credential.c2pa").Open()) WriteAll(manifest, CreateManifestStore());
                using (Stream keep = archive.CreateEntry("keep.txt").Open()) WriteAll(keep, Encoding.UTF8.GetBytes("keep"));
            }
            package = AddCentralDirectoryComment(stream.ToArray(), "keep.txt", rawComment);
        }
        package = AddEntryExtraFields(package, "keep.txt", Array.Empty<byte>(), unicodeExtra);
        int sourceCentralHeader = FindSignature(package, 0x02014B50u, "keep.txt");
        WriteLittleEndian16(package, sourceCentralHeader + 8,
            (ushort)(BitConverter.ToUInt16(package, sourceCentralHeader + 8) & ~0x0800));
        int sourceLocalHeader = checked((int)BitConverter.ToUInt32(package, sourceCentralHeader + 42));
        WriteLittleEndian16(package, sourceLocalHeader + 6,
            (ushort)(BitConverter.ToUInt16(package, sourceLocalHeader + 6) & ~0x0800));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "fixture.zip");
        int centralHeader = FindSignature(result.ToArray(), 0x02014B50u, "keep.txt");

        Assert.Equal("intended Ω", Encoding.UTF8.GetString(ReadCentralDirectoryComment(result.ToArray(), centralHeader)));
        Assert.Empty(ReadCentralExtraField(result.ToArray(), centralHeader));
        Assert.NotEqual(0, BitConverter.ToUInt16(result.ToArray(), centralHeader + 8) & 0x0800);
    }
}
