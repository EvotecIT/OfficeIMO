namespace OfficeIMO.Email.Store.Tests;

internal static class PstTestFileBuilder {
    private const int BbtOffset = 1024;
    private const int NbtOffset = 1536;

    internal static byte[] Create(bool ost = false) {
        var nodes = new[] {
            new TestNode(0x21, 0, CreatePropertyContext((0x3001, "Test Store"))),
            new TestNode(0x122, 0x122, CreatePropertyContext((0x3001, "Root"))),
            new TestNode(0x8022, 0x122, CreatePropertyContext((0x3001, "Inbox"))),
            new TestNode(0x8004, 0x8022, CreatePropertyContext(
                (0x001A, "IPM.Note"),
                (0x0037, "Synthetic PST message"),
                (0x1000, "Body from the PST property context")))
        };

        int blockOffset = 2048;
        for (int index = 0; index < nodes.Length; index++) {
            nodes[index].Bid = (ulong)(0x100 + index * 4);
            nodes[index].Offset = blockOffset;
            blockOffset += Align64(nodes[index].Data.Length + 16);
        }
        var file = new byte[blockOffset];
        WriteHeader(file, ost);
        WriteBbt(file, nodes);
        WriteNbt(file, nodes);
        foreach (TestNode node in nodes) Buffer.BlockCopy(node.Data, 0, file, node.Offset, node.Data.Length);
        return file;
    }

    private static void WriteHeader(byte[] file, bool ost) {
        WriteUInt32(file, 0, 0x4E444221);
        file[8] = 0x53;
        file[9] = ost ? (byte)0x4F : (byte)0x4D;
        WriteUInt16(file, 10, 23);
        WriteUInt16(file, 12, 19);
        file[14] = 1;
        file[15] = 1;
        WriteUInt64(file, 216, 0x200);
        WriteUInt64(file, 224, NbtOffset);
        WriteUInt64(file, 232, 0x204);
        WriteUInt64(file, 240, BbtOffset);
        file[512] = 0x80;
        file[513] = 0;
    }

    private static void WriteBbt(byte[] file, IReadOnlyList<TestNode> nodes) {
        for (int index = 0; index < nodes.Count; index++) {
            int offset = BbtOffset + index * 24;
            TestNode node = nodes[index];
            WriteUInt64(file, offset, node.Bid);
            WriteUInt64(file, offset + 8, node.Offset);
            WriteUInt16(file, offset + 16, node.Data.Length);
            WriteUInt16(file, offset + 18, 1);
        }
        WriteBTreePageMetadata(file, BbtOffset, nodes.Count, 24, 0x80);
    }

    private static void WriteNbt(byte[] file, IReadOnlyList<TestNode> nodes) {
        for (int index = 0; index < nodes.Count; index++) {
            int offset = NbtOffset + index * 32;
            TestNode node = nodes[index];
            WriteUInt32(file, offset, node.Nid);
            WriteUInt64(file, offset + 8, node.Bid);
            WriteUInt64(file, offset + 16, 0);
            WriteUInt32(file, offset + 24, node.ParentNid);
        }
        WriteBTreePageMetadata(file, NbtOffset, nodes.Count, 32, 0x81);
    }

    private static void WriteBTreePageMetadata(byte[] file, int pageOffset, int count, int entrySize, byte type) {
        file[pageOffset + 488] = checked((byte)count);
        file[pageOffset + 489] = checked((byte)(488 / entrySize));
        file[pageOffset + 490] = checked((byte)entrySize);
        file[pageOffset + 491] = 0;
        file[pageOffset + 496] = type;
        file[pageOffset + 497] = type;
    }

    private static byte[] CreatePropertyContext(params (ushort Id, string Value)[] strings) {
        int allocationCount = 2 + strings.Length;
        var allocations = new List<byte[]> {
            new byte[] { 0xB5, 0x02, 0x06, 0x00, 0x40, 0x00, 0x00, 0x00 }
        };
        var records = new byte[strings.Length * 8];
        for (int index = 0; index < strings.Length; index++) {
            int recordOffset = index * 8;
            WriteUInt16(records, recordOffset, strings[index].Id);
            WriteUInt16(records, recordOffset + 2, 0x001F);
            uint hid = checked((uint)((3 + index) << 5));
            WriteUInt32(records, recordOffset + 4, hid);
        }
        allocations.Add(records);
        foreach ((ushort _, string value) in strings) allocations.Add(Encoding.Unicode.GetBytes(value + "\0"));

        int dataEnd = 12 + allocations.Sum(item => item.Length);
        int mapOffset = (dataEnd + 1) & ~1;
        int mapLength = 4 + (allocationCount + 1) * 2;
        var block = new byte[mapOffset + mapLength];
        WriteUInt16(block, 0, mapOffset);
        block[2] = 0xEC;
        block[3] = 0xBC;
        WriteUInt32(block, 4, 0x20);

        int cursor = 12;
        var boundaries = new List<int> { cursor };
        foreach (byte[] allocation in allocations) {
            Buffer.BlockCopy(allocation, 0, block, cursor, allocation.Length);
            cursor += allocation.Length;
            boundaries.Add(cursor);
        }
        WriteUInt16(block, mapOffset, allocationCount);
        WriteUInt16(block, mapOffset + 2, 0);
        for (int index = 0; index < boundaries.Count; index++) {
            WriteUInt16(block, mapOffset + 4 + index * 2, boundaries[index]);
        }
        return block;
    }

    private static int Align64(int value) => (value + 63) & ~63;

    private static void WriteUInt16(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
    }

    private static void WriteUInt32(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
        bytes[offset + 3] = (byte)(value >> 24);
    }

    private static void WriteUInt64(byte[] bytes, int offset, long value) => WriteUInt64(bytes, offset, (ulong)value);

    private static void WriteUInt64(byte[] bytes, int offset, ulong value) {
        WriteUInt32(bytes, offset, (uint)value);
        WriteUInt32(bytes, offset + 4, (uint)(value >> 32));
    }

    private sealed class TestNode {
        internal TestNode(uint nid, uint parentNid, byte[] data) {
            Nid = nid;
            ParentNid = parentNid;
            Data = data;
        }

        internal uint Nid { get; }
        internal uint ParentNid { get; }
        internal byte[] Data { get; }
        internal ulong Bid { get; set; }
        internal int Offset { get; set; }
    }
}
