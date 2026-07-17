using System.Threading;

namespace OfficeIMO.Email.Store.Tests;

public sealed class PstHeapTests {
    [Fact]
    public void ResolvesVersion36ThirteenBitHeapBlockIndexes() {
        using var stream = new MemoryStream(PstTestFileBuilder.Create(ost: true, fourK: true));
        PstHeader header = PstHeader.Read(stream, EmailStoreFormat.Ost);
        var ndb = new PstNdbReader(stream, header, EmailStoreReaderOptions.Default,
            CancellationToken.None);
        byte[] expected = new byte[] { 1, 3, 5, 7, 9 };
        var dataTree = new PstDataTree(new[] {
            CreateHeapPage(Array.Empty<byte>(), firstPage: true),
            CreateHeapPage(expected, firstPage: false)
        }, expected.LongLength);
        var heap = new PstHeap(dataTree, new Dictionary<uint, PstSubnodeReference>(), ndb,
            EmailStoreReaderOptions.Default, CancellationToken.None);

        byte[] actual = heap.GetAllocation(0x00080020);

        Assert.Equal(expected, actual);
    }

    [Fact]
    public void ResolvesVersion36FourteenBitHeapAllocationIndexes() {
        using var stream = new MemoryStream(PstTestFileBuilder.Create(ost: true, fourK: true));
        PstHeader header = PstHeader.Read(stream, EmailStoreFormat.Ost);
        var ndb = new PstNdbReader(stream, header, EmailStoreReaderOptions.Default,
            CancellationToken.None);
        byte[] expected = new byte[] { 2, 4, 6, 8 };
        var dataTree = new PstDataTree(new[] {
            CreateHeapPageWithAllocations(2049, 2048, expected)
        }, expected.LongLength);
        var heap = new PstHeap(dataTree, new Dictionary<uint, PstSubnodeReference>(), ndb,
            EmailStoreReaderOptions.Default, CancellationToken.None);

        byte[] actual = heap.GetAllocation(0x00010020);

        Assert.Equal(expected, actual);
    }

    private static byte[] CreateHeapPage(byte[] allocation, bool firstPage, int minimumLength = 0) {
        int headerSize = firstPage ? 12 : 8;
        int mapOffset = (headerSize + allocation.Length + 1) & ~1;
        var block = new byte[Math.Max(mapOffset + 8, minimumLength)];
        WriteUInt16(block, 0, mapOffset);
        if (firstPage) {
            block[2] = 0xEC;
            block[3] = 0xBC;
        }
        Buffer.BlockCopy(allocation, 0, block, headerSize, allocation.Length);
        WriteUInt16(block, mapOffset, 1);
        WriteUInt16(block, mapOffset + 4, headerSize);
        WriteUInt16(block, mapOffset + 6, headerSize + allocation.Length);
        return block;
    }

    private static byte[] CreateHeapPageWithAllocations(
        int allocationCount, int targetIndex, byte[] target) {
        const int headerSize = 12;
        int dataLength = allocationCount - 1 + target.Length;
        int mapOffset = (headerSize + dataLength + 1) & ~1;
        var block = new byte[mapOffset + 4 + (allocationCount + 1) * 2];
        WriteUInt16(block, 0, mapOffset);
        block[2] = 0xEC;
        block[3] = 0xBC;
        int cursor = headerSize;
        WriteUInt16(block, mapOffset, allocationCount);
        for (int index = 0; index < allocationCount; index++) {
            WriteUInt16(block, mapOffset + 4 + index * 2, cursor);
            if (index == targetIndex) {
                Buffer.BlockCopy(target, 0, block, cursor, target.Length);
                cursor += target.Length;
            } else {
                block[cursor++] = 0xCC;
            }
        }
        WriteUInt16(block, mapOffset + 4 + allocationCount * 2, cursor);
        return block;
    }

    private static void WriteUInt16(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
    }
}
