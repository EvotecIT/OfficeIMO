using System.IO.Compression;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

public sealed partial class ZipExtractionTests {
    [Theory]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(true, false, true)]
    public void Zip64DescriptorWidthComesFromEntryExtraFields(bool signature, bool local64, bool central64) {
        byte[] valid = EmptyDescriptorZip(signature, local64, central64, 0);
        using (var fixture = new ArchiveFixtures()) Assert.Single(ArchiveExtractor.Extract(fixture.Package(valid), fixture.Destination).Files);
        // Its first 12 bytes resemble a valid classic descriptor, but its ZIP64 expanded size contradicts the directory.
        byte[] invalid = EmptyDescriptorZip(signature, local64, central64, 1);
        using (var fixture = new ArchiveFixtures()) {
            Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(invalid), fixture.Destination,
                new ArchiveExtractionOptions { IncludeFile = _ => false }));
            Assert.False(Directory.Exists(fixture.Destination));
        }
        using var input = new MemoryStream(invalid, writable: false);
        using var archive = new ZipArchive(input, ZipArchiveMode.Read);
        var entry = Assert.Single(archive.Entries);
        var metadata = OfficeIMO.Provenance.OfficeProvenanceZip.GetInputEntryMetadata(invalid, archive)[entry];
        Assert.Throws<InvalidDataException>(() => OfficeArchiveSafety.OpenEntryPayload(invalid, entry, metadata));
    }

    private static byte[] EmptyDescriptorZip(bool signature, bool local64, bool central64, ulong descriptorExpanded) {
        using var source = new MemoryStream();
        using var writer = new BinaryWriter(source, Encoding.ASCII, leaveOpen: true);
        byte[] name = Encoding.ASCII.GetBytes("empty.admx");
        writer.Write(0x04034b50U); writer.Write((ushort)45); writer.Write((ushort)8); writer.Write((ushort)0);
        writer.Write(0U); writer.Write(0U); writer.Write(local64 ? uint.MaxValue : 0U); writer.Write(local64 ? uint.MaxValue : 0U);
        writer.Write((ushort)name.Length); writer.Write((ushort)(local64 ? 20 : 0)); writer.Write(name);
        if (local64) WriteExtra();
        if (signature) writer.Write(0x08074b50U);
        writer.Write(0U); writer.Write(0UL); writer.Write(descriptorExpanded);
        uint central = (uint)source.Position;
        writer.Write(0x02014b50U); writer.Write((ushort)45); writer.Write((ushort)45); writer.Write((ushort)8); writer.Write((ushort)0);
        writer.Write(0U); writer.Write(0U); writer.Write(central64 ? uint.MaxValue : 0U); writer.Write(central64 ? uint.MaxValue : 0U);
        writer.Write((ushort)name.Length); writer.Write((ushort)(central64 ? 20 : 0)); writer.Write((ushort)0);
        writer.Write((ushort)0); writer.Write((ushort)0); writer.Write(0U); writer.Write(0U); writer.Write(name);
        if (central64) WriteExtra();
        uint size = (uint)source.Position - central;
        writer.Write(0x06054b50U); writer.Write((ushort)0); writer.Write((ushort)0); writer.Write((ushort)1); writer.Write((ushort)1);
        writer.Write(size); writer.Write(central); writer.Write((ushort)0);
        return source.ToArray();
        void WriteExtra() { writer.Write((ushort)1); writer.Write((ushort)16); writer.Write(0UL); writer.Write(0UL); }
    }

    [Fact]
    public void BytePreflightRejectsASecondZip64DirectoryBeforeApplyingEntryLimits() {
        byte[] classic = Zip(("index.html", "<p>Hello</p>", CompressionLevel.NoCompression), ("other.txt", "B", CompressionLevel.NoCompression));
        int end = classic.Length - 22, central = (int)BitConverter.ToUInt32(classic, end + 16);
        int firstLength = 46 + BitConverter.ToUInt16(classic, central + 28) + BitConverter.ToUInt16(classic, central + 30) + BitConverter.ToUInt16(classic, central + 32);
        using var bytes = new MemoryStream();
        using var writer = new BinaryWriter(bytes, Encoding.ASCII, leaveOpen: true);
        writer.Write(classic, 0, end); writer.Write(0U);
        ulong alternate = (ulong)bytes.Position;
        writer.Write(classic, central, firstLength);
        ulong record = (ulong)bytes.Position;
        writer.Write(0x06064b50U); writer.Write(44UL); writer.Write((ushort)45); writer.Write((ushort)45); writer.Write(0U); writer.Write(0U);
        writer.Write(1UL); writer.Write(1UL); writer.Write((ulong)firstLength); writer.Write(alternate);
        writer.Write(0x07064b50U); writer.Write(0U); writer.Write(record); writer.Write(1U); writer.Write(classic, end, 22);
        byte[] zip = bytes.ToArray();
        using var input = new MemoryStream(zip);
        using var archive = new ZipArchive(input, ZipArchiveMode.Read);
        Assert.Equal(2, archive.Entries.Count); // The runtime follows the unmasked classic directory.
        Assert.False(OfficeArchiveSafety.ScanZipCentralDirectory(zip, 1).IsValid);
        input.Position = 0;
        Assert.False(OfficeArchiveSafety.ScanZipCentralDirectory(input, input.Length, 1).IsValid);
    }
}
