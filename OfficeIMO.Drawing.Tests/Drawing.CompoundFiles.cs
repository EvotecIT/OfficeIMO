using OfficeIMO.Drawing.Internal;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingCompoundFileTests {
    private const int SectorSize = 4096;
    private const uint FreeSect = 0xffffffff;
    private const uint EndOfChain = 0xfffffffe;
    private const uint FatSect = 0xfffffffd;

    [Fact]
    public void ReaderFindsVersion4SectorsAfterThePaddedHeaderSector() {
        byte[] compound = CreateVersion4RootOnlyCompoundFile();

        bool success = OfficeCompoundFileReader.TryRead(compound, out OfficeCompoundFile? file, out string? error);

        Assert.True(success, error);
        Assert.NotNull(file);
        Assert.Empty(file!.Streams);
        Assert.Contains(file.Entries, entry => entry.IsStorage && entry.Name == "Root Entry");
    }

    [Fact]
    public void RewritePreservesStreamsEmptyStoragesAndDirectoryMetadata() {
        Guid rootClassId = new Guid("64818D10-4F9B-11CF-86EA-00AA00B929E8");
        Guid storageClassId = new Guid("00020820-0000-0000-C000-000000000046");
        const ulong created = 132537600000000000;
        const ulong modified = 132537636000000000;
        var streams = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase) {
            ["PowerPoint Document"] = new byte[] { 1, 2, 3 },
            ["ObjectPool/Item/Contents"] = new byte[] { 4, 5, 6 }
        };
        var source = new OfficeCompoundFile(streams, new[] {
            new OfficeCompoundFileEntry("PowerPoint Document", "PowerPoint Document", 2, 3),
            new OfficeCompoundFileEntry("ObjectPool", "ObjectPool", 1, 0),
            new OfficeCompoundFileEntry("Item", "ObjectPool/Item", 1, 0, classId: storageClassId,
                stateBits: 7, creationTime: created, modifiedTime: modified),
            new OfficeCompoundFileEntry("Contents", "ObjectPool/Item/Contents", 2, 3),
            new OfficeCompoundFileEntry("EmptyStorage", "EmptyStorage", 1, 0, stateBits: 9)
        }, new OfficeCompoundFileEntry("Root Entry", "Root Entry", 5, 0, classId: rootClassId));

        byte[] rewritten = OfficeCompoundFileWriter.Rewrite(source,
            new Dictionary<string, byte[]> { ["PowerPoint Document"] = new byte[] { 9, 8, 7, 6 } });
        bool success = OfficeCompoundFileReader.TryRead(rewritten, out OfficeCompoundFile? roundTrip,
            out string? error);

        Assert.True(success, error);
        Assert.NotNull(roundTrip);
        Assert.Equal(new byte[] { 9, 8, 7, 6 }, roundTrip!.Streams["PowerPoint Document"]);
        Assert.Equal(new byte[] { 4, 5, 6 }, roundTrip.Streams["ObjectPool/Item/Contents"]);
        Assert.Equal(rootClassId, roundTrip.RootEntry.ClassId);
        OfficeCompoundFileEntry storage = Assert.Single(roundTrip.Entries,
            entry => entry.Path == "ObjectPool/Item");
        Assert.Equal(storageClassId, storage.ClassId);
        Assert.Equal(7U, storage.StateBits);
        Assert.Equal(created, storage.CreationTime);
        Assert.Equal(modified, storage.ModifiedTime);
        Assert.Contains(roundTrip.Entries, entry => entry.Path == "EmptyStorage" && entry.StateBits == 9);
    }

    [Fact]
    public void DocumentDetectionHonorsCancellationDuringDirectoryInspection() {
        byte[] compound = CreateVersion4RootOnlyCompoundFile();
        using var cancellation = new CancellationTokenSource();
        using var stream = new CancelAfterReadStream(compound, 2,
            cancellation.Cancel);

        void Detect() => OfficeCompoundDocumentDetector.Detect(stream,
            compound.LongLength, 65536, cancellation.Token, out _);
        Assert.Throws<OperationCanceledException>(Detect);

        Assert.Equal(2, stream.ReadCount);
        Assert.Equal(0, stream.Position);
    }

    private static byte[] CreateVersion4RootOnlyCompoundFile() {
        byte[] compound = new byte[SectorSize * 3];
        byte[] signature = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
        Buffer.BlockCopy(signature, 0, compound, 0, signature.Length);
        WriteUInt16(compound, 24, 0x003e);
        WriteUInt16(compound, 26, 0x0004);
        WriteUInt16(compound, 28, 0xfffe);
        WriteUInt16(compound, 30, 0x000c);
        WriteUInt16(compound, 32, 0x0006);
        WriteUInt32(compound, 40, 1);
        WriteUInt32(compound, 44, 1);
        WriteUInt32(compound, 48, 0);
        WriteUInt32(compound, 56, 4096);
        WriteUInt32(compound, 60, EndOfChain);
        WriteUInt32(compound, 68, EndOfChain);
        for (int index = 0; index < 109; index++) {
            WriteUInt32(compound, 76 + index * 4, index == 0 ? 1U : FreeSect);
        }

        int directoryOffset = SectorSize;
        byte[] rootName = Encoding.Unicode.GetBytes("Root Entry\0");
        Buffer.BlockCopy(rootName, 0, compound, directoryOffset, rootName.Length);
        WriteUInt16(compound, directoryOffset + 64, checked((ushort)rootName.Length));
        compound[directoryOffset + 66] = 5;
        compound[directoryOffset + 67] = 1;
        WriteUInt32(compound, directoryOffset + 68, FreeSect);
        WriteUInt32(compound, directoryOffset + 72, FreeSect);
        WriteUInt32(compound, directoryOffset + 76, FreeSect);
        WriteUInt32(compound, directoryOffset + 116, EndOfChain);

        int fatOffset = SectorSize * 2;
        WriteUInt32(compound, fatOffset, EndOfChain);
        WriteUInt32(compound, fatOffset + 4, FatSect);
        for (int index = 2; index < SectorSize / 4; index++) {
            WriteUInt32(compound, fatOffset + index * 4, FreeSect);
        }

        return compound;
    }

    private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
    }

    private static void WriteUInt32(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
        bytes[offset + 3] = (byte)(value >> 24);
    }

    private sealed class CancelAfterReadStream : MemoryStream {
        private readonly int _cancelAfterRead;
        private readonly Action _cancel;

        public CancelAfterReadStream(byte[] bytes, int cancelAfterRead,
            Action cancel) : base(bytes, writable: false) {
            _cancelAfterRead = cancelAfterRead;
            _cancel = cancel;
        }

        public int ReadCount { get; private set; }

        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            ReadCount++;
            if (ReadCount == _cancelAfterRead) _cancel();
            return read;
        }
    }
}
