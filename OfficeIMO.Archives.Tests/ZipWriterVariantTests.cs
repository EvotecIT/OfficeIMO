using System.IO.Compression;
using System.Text;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

// Archives that other writers produce and that every reader resolves to the same members.
public sealed partial class ZipExtractionTests {
    [Fact]
    public void BytesAppendedAfterTheEndRecordAreIgnored() {
        using var fixture = new ArchiveFixtures();
        byte[] zip = Zip(("a.admx", "A", CompressionLevel.NoCompression)).Concat(new byte[] { 0x0a }).ToArray();

        Assert.True(OfficeArchiveSafety.ScanZipCentralDirectory(zip, 10).IsValid);
        Assert.Equal("A", File.ReadAllText(Path.Combine(fixture.Destination, Assert.Single(
            ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination).Files).RelativePath)));
    }

    [Fact]
    public void AppendedBytesHoldingAnotherEndRecordAreRejected() {
        byte[] zip = Zip(("a.admx", "A", CompressionLevel.NoCompression));
        byte[] crafted = zip.Concat(new byte[] { 0x0a }).Concat(zip.Skip(zip.Length - 22)).Concat(new byte[] { 0x0a }).ToArray();

        Assert.False(OfficeArchiveSafety.ScanZipCentralDirectory(crafted, 10).IsValid);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DescriptorsWithEightByteSizesAreAcceptedForSmallMembers(bool signed) {
        using var fixture = new ArchiveFixtures();
        byte[] zip = StoredMember("a.txt", "Hello", localFlags: 8, centralFlags: 8, descriptor: Descriptor(signed, wide: true, "Hello"));

        var result = ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination);
        Assert.Equal("Hello", File.ReadAllText(Path.Combine(fixture.Destination, Assert.Single(result.Files).RelativePath)));
    }

    [Fact]
    public void EightByteDescriptorSizesMustAgreeWithTheCentralRecord() {
        using var fixture = new ArchiveFixtures();
        byte[] descriptor = Descriptor(signed: true, wide: true, "Hello");
        descriptor[16]++; // expanded size

        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(StoredMember("a.txt", "Hello", localFlags: 8, centralFlags: 8, descriptor: descriptor)), fixture.Destination));
    }

    [Theory]
    [InlineData("a.txt", true)]
    [InlineData("\u00e4.txt", false)]
    public void AUtf8NameFlagOnlyInTheCentralRecordIsAcceptedForAsciiNames(string name, bool accepted) {
        using var fixture = new ArchiveFixtures();
        byte[] zip = StoredMember(name, "Hello", localFlags: 0, centralFlags: 0x0800, descriptor: Array.Empty<byte>());

        if (accepted) {
            var result = ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination);
            Assert.Equal("Hello", File.ReadAllText(Path.Combine(fixture.Destination, Assert.Single(result.Files).RelativePath)));
        } else {
            Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination));
        }
    }

    // One stored member. The local header carries the checksum and sizes unless its descriptor flag is set.
    private static byte[] StoredMember(string name, string text, ushort localFlags, ushort centralFlags, byte[] descriptor) {
        byte[] nameBytes = Encoding.UTF8.GetBytes(name);
        byte[] data = Encoding.ASCII.GetBytes(text);
        uint crc = Crc32(data);
        bool localHasSizes = (localFlags & 8) == 0;
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream);
        writer.Write(0x04034b50U); writer.Write((ushort)20); writer.Write(localFlags); writer.Write((ushort)0);
        writer.Write(0U);
        writer.Write(localHasSizes ? crc : 0U);
        writer.Write(localHasSizes ? (uint)data.Length : 0U);
        writer.Write(localHasSizes ? (uint)data.Length : 0U);
        writer.Write((ushort)nameBytes.Length); writer.Write((ushort)0); writer.Write(nameBytes); writer.Write(data);
        writer.Write(descriptor);
        long central = stream.Position;
        writer.Write(0x02014b50U); writer.Write((ushort)20); writer.Write((ushort)20); writer.Write(centralFlags); writer.Write((ushort)0);
        writer.Write(0U); writer.Write(crc); writer.Write((uint)data.Length); writer.Write((uint)data.Length);
        writer.Write((ushort)nameBytes.Length); writer.Write((ushort)0); writer.Write((ushort)0); writer.Write((ushort)0); writer.Write((ushort)0);
        writer.Write(0U); writer.Write(0U); writer.Write(nameBytes);
        long end = stream.Position;
        writer.Write(0x06054b50U); writer.Write((ushort)0); writer.Write((ushort)0); writer.Write((ushort)1); writer.Write((ushort)1);
        writer.Write((uint)(end - central)); writer.Write((uint)central); writer.Write((ushort)0);
        writer.Flush();
        return stream.ToArray();
    }

    private static byte[] Descriptor(bool signed, bool wide, string text) {
        byte[] data = Encoding.ASCII.GetBytes(text);
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream);
        if (signed) writer.Write(0x08074b50U);
        writer.Write(Crc32(data));
        if (wide) { writer.Write((ulong)data.Length); writer.Write((ulong)data.Length); } else { writer.Write((uint)data.Length); writer.Write((uint)data.Length); }
        writer.Flush();
        return stream.ToArray();
    }

    private static uint Crc32(byte[] data) {
        uint crc = 0xffffffffU;
        foreach (byte value in data) {
            crc ^= value;
            for (int bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xedb88320U : crc >> 1;
        }
        return ~crc;
    }
}
