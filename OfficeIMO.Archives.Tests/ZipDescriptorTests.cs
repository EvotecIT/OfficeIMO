using System.IO.Compression;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

public sealed partial class ZipExtractionTests {
    [Theory]
    [InlineData(0x0010, false)]
    [InlineData(0x0080, false)]
    [InlineData(0x0100, false)]
    [InlineData(0x0200, false)]
    [InlineData(0x0400, false)]
    [InlineData(0x1000, false)]
    [InlineData(0x4000, false)]
    [InlineData(0x8000, false)]
    [InlineData(0x0002, true)]
    [InlineData(0x0004, true)]
    public void UnsupportedFlagsFailBeforeSelectionInBothPayloadReaders(int flag, bool stored) {
        byte[] zip = PatchHeaders(stored ? StoredMemberZip() : Zip(("Example.admx", "A", CompressionLevel.Optimal)), flagsToSet: (ushort)flag);
        using var fixture = new ArchiveFixtures();
        int calls = 0;
        Assert.Throws<NotSupportedException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = _ => { calls++; return false; } }));
        Assert.Equal(0, calls);
        Assert.False(Directory.Exists(fixture.Destination));
        using var input = new MemoryStream(zip);
        using var archive = new ZipArchive(input, ZipArchiveMode.Read);
        var entry = Assert.Single(archive.Entries);
        var metadata = OfficeIMO.Provenance.OfficeProvenanceZip.GetInputEntryMetadata(zip, archive)[entry];
        Assert.Throws<InvalidDataException>(() => OfficeArchiveSafety.OpenEntryPayload(zip, entry, metadata));
    }

    [Theory]
    [InlineData(2)]
    [InlineData(4)]
    [InlineData(6)]
    public void DeflateCompressionLevelFlagsRemainSupported(int flags) {
        using var fixture = new ArchiveFixtures();
        byte[] zip = PatchHeaders(Zip(("Example.admx", "A", CompressionLevel.Optimal)), flagsToSet: (ushort)flags);
        Assert.Single(ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination).Files);
    }

    [Fact]
    public void StoredMembersAcceptUtf8NamesWithoutDeflateLevelHints() {
        using var fixture = new ArchiveFixtures();
        Assert.Single(ArchiveExtractor.Extract(fixture.Package(PatchHeaders(StoredMemberZip(), flagsToSet: 0x0800)), fixture.Destination).Files);
    }

    private static byte[] StoredMemberZip() {
        byte[] original = Zip(("Example.admx", "A", CompressionLevel.Optimal));
        int payload = 30 + BitConverter.ToUInt16(original, 26) + BitConverter.ToUInt16(original, 28);
        int central = (int)BitConverter.ToUInt32(original, original.Length - 6);
        byte[] zip = original.Take(payload).Concat(new[] { (byte)'A' }).Concat(original.Skip(central)).ToArray();
        central = payload + 1;
        WriteUInt16(zip, 8, 0); WriteUInt16(zip, central + 10, 0);
        WriteUInt32(zip, 18, 1); WriteUInt32(zip, central + 20, 1); WriteUInt32(zip, zip.Length - 6, (uint)central);
        return zip;
    }

    [Fact]
    public void EndRecordNumericFieldsAreNotMistakenForCommentSignatures() {
        // A virtual read-only archive avoids allocating its 101 MB central-directory comments.
        using var source = new LargeCommentDirectoryStream();
        var scan = OfficeArchiveSafety.ScanZipCentralDirectory(source, source.Length, 2000);
        Assert.True(scan.IsValid, scan.Error);
        Assert.Equal(1541, scan.EntryCount);
        Assert.Equal(0, source.Position);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EndSignaturesOverlappingTheCommentBoundaryRemainAmbiguous(bool oneHeaderByte) {
        byte[] original = Zip(("Example.admx", "A", CompressionLevel.Optimal));
        byte[] comment = new byte[oneHeaderByte ? 0x5000 : 0x4b50];
        if (oneHeaderByte) { comment[0] = 0x4b; comment[1] = 5; comment[2] = 6; }
        else { comment[0] = 5; comment[1] = 6; }
        WriteUInt16(original, original.Length - 2, (ushort)comment.Length);
        byte[] zip = original.Concat(comment).ToArray();
        Assert.False(OfficeArchiveSafety.ScanZipCentralDirectory(zip, 10).IsValid);
        using var fixture = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination));
    }

    private sealed class LargeCommentDirectoryStream : Stream {
        private const int DirectorySize = 0x06054b50, RecordLength = 65582, LocalLength = 31;
        private readonly byte[] _local = new byte[LocalLength], _full = Header(ushort.MaxValue), _last = Header((ushort)(DirectorySize % RecordLength - 47)), _end = new byte[22];
        internal LargeCommentDirectoryStream() {
            WriteUInt32(_local, 0, 0x04034b50); WriteUInt16(_local, 26, 1); _local[30] = (byte)'A';
            WriteUInt32(_end, 0, 0x06054b50); WriteUInt16(_end, 8, 1541); WriteUInt16(_end, 10, 1541);
            WriteUInt32(_end, 12, DirectorySize); WriteUInt32(_end, 16, LocalLength);
        }
        private static byte[] Header(ushort comment) {
            var bytes = new byte[47]; WriteUInt32(bytes, 0, 0x02014b50); WriteUInt16(bytes, 28, 1); WriteUInt16(bytes, 32, comment); bytes[46] = (byte)'A'; return bytes;
        }
        public override int Read(byte[] buffer, int offset, int count) {
            int total = 0;
            while (count > 0 && Position < Length) {
                byte[]? data; int within, available;
                if (Position < LocalLength) { data = _local; within = (int)Position; available = LocalLength - within; }
                else if (Position >= LocalLength + DirectorySize) { data = _end; within = (int)(Position - LocalLength - DirectorySize); available = data.Length - within; }
                else {
                    int relative = (int)(Position - LocalLength), record = relative / RecordLength;
                    within = relative % RecordLength; data = record == DirectorySize / RecordLength ? _last : _full;
                    int length = record == DirectorySize / RecordLength ? DirectorySize % RecordLength : RecordLength;
                    available = within < data.Length ? data.Length - within : length - within;
                    if (within >= data.Length) data = null;
                }
                int read = Math.Min(count, available);
                if (data == null) Array.Clear(buffer, offset, read); else Array.Copy(data, within, buffer, offset, read);
                Position += read; offset += read; count -= read; total += read;
            }
            return total;
        }
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => LocalLength + DirectorySize + 22L;
        public override long Position { get; set; }
        public override long Seek(long offset, SeekOrigin origin) => Position = (origin == SeekOrigin.Begin ? 0 : origin == SeekOrigin.Current ? Position : Length) + offset;
        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

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
