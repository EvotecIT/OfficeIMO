namespace OfficeIMO.Email.Store.Tests;

internal static class PstTestFileBuilder {
    private const int BbtOffset = 1024;
    private const int NbtOffset = 1536;

    internal static byte[] Create(bool ost = false, bool ansi = false, byte cryptMethod = 0,
        bool fourK = false, bool compressBlocks = false, bool includeEmbeddedMessage = false,
        byte[]? attachmentContent = null, uint? storePasswordChecksum = null,
        bool corruptPageCrc = false, bool corruptBlockCrc = false,
        bool includeFixedNidSearchFolder = false, uint? inboxParentNid = null) {
        if (fourK && (!ost || ansi)) throw new ArgumentException("4K test stores must use the Unicode OST variant.");
        if (compressBlocks && !fourK) throw new ArgumentException("Only 4K test blocks can be compressed.");
        if (compressBlocks && cryptMethod != 0) throw new ArgumentException("The fixture does not combine compression and encoding.");
        if (includeEmbeddedMessage && attachmentContent != null) {
            throw new ArgumentException("The compact fixture supports one attachment shape at a time.");
        }
        var nodeList = new List<TestNode> {
            new TestNode(0x21, 0, storePasswordChecksum.HasValue
                ? CreateStorePropertyContext(storePasswordChecksum.Value)
                : CreatePropertyContext((0x3001, "Test Store"))),
            new TestNode(0x122, 0x122, CreatePropertyContext((0x3001, "Root"))),
            new TestNode(0x8022, inboxParentNid ?? 0x122, CreatePropertyContext((0x3001, "Inbox"))),
            new TestNode(0x8004, 0x8022, CreatePropertyContext(
                (0x001A, "IPM.Note"),
                (0x0037, "Synthetic PST message"),
                (0x1000, "Body from the PST property context")))
        };
        if (includeFixedNidSearchFolder) {
            nodeList.Add(new TestNode(0x2223, 0x122, CreatePropertyContext(
                (0x3001, "External fixed-NID search"),
                (0x3613, "IPF.Note"))));
        }
        TestNode[] nodes = nodeList.ToArray();

        var blocks = nodes.Select(node => node.DataBlock).ToList();
        if (includeEmbeddedMessage) AddEmbeddedMessageBlocks(nodes[3], blocks, ansi);
        if (attachmentContent != null) AddByValueAttachmentBlocks(nodes[3], blocks, attachmentContent, ansi);

        int bbtOffset = fourK ? 4096 : BbtOffset;
        int nbtOffset = fourK ? 8192 : NbtOffset;
        int blockOffset = fourK ? 12288 : 2048;
        int blockTrailerSize = ansi ? 12 : fourK ? 24 : 16;
        int blockAlignment = fourK ? 512 : 64;
        for (int index = 0; index < nodes.Length; index++) {
            nodes[index].DataBlock.Bid = (ulong)(0x100 + index * 4);
        }
        foreach (TestBlock block in blocks) {
            block.StoredData = compressBlocks ? Compress(block.Data) : block.Data.ToArray();
            block.Offset = blockOffset;
            block.AllocationLength = Align(block.StoredData.Length + blockTrailerSize, blockAlignment);
            blockOffset += block.AllocationLength;
        }
        var file = new byte[blockOffset];
        WriteHeader(file, ost, ansi, fourK, cryptMethod, bbtOffset, nbtOffset);
        WriteBbt(file, blocks, ansi, fourK, bbtOffset);
        WriteNbt(file, nodes, ansi, fourK, nbtOffset);
        foreach (TestBlock block in blocks) {
            byte[] data = block.StoredData.ToArray();
            if (cryptMethod != 0 && (block.Bid & 0x02) == 0) PstCrypt.Decode(data, cryptMethod, block.Bid);
            Buffer.BlockCopy(data, 0, file, block.Offset, data.Length);
            WriteBlockTrailer(file, block, data, ansi, fourK, blockTrailerSize);
        }
        if (corruptPageCrc) {
            int trailerOffset = nbtOffset + (ansi ? 500 : fourK ? 4072 : 496);
            int crcOffset = trailerOffset + (ansi ? 8 : 4);
            file[crcOffset] ^= 0x01;
        }
        if (corruptBlockCrc) {
            TestBlock messageBlock = nodes[3].DataBlock;
            int trailerOffset = messageBlock.Offset + messageBlock.AllocationLength - blockTrailerSize;
            file[trailerOffset + 4] ^= 0x01;
        }
        return file;
    }

    private static void WriteHeader(byte[] file, bool ost, bool ansi, bool fourK, byte cryptMethod,
        int bbtOffset, int nbtOffset) {
        WriteUInt32(file, 0, 0x4E444221);
        file[8] = 0x53;
        file[9] = ost ? (byte)0x4F : (byte)0x4D;
        WriteUInt16(file, 10, ansi ? 15 : fourK ? 36 : 23);
        WriteUInt16(file, 12, ansi ? 15 : 19);
        file[14] = 1;
        file[15] = 1;
        if (ansi) {
            WriteUInt32(file, 184, 0x200);
            WriteUInt32(file, 188, checked((uint)nbtOffset));
            WriteUInt32(file, 192, 0x204);
            WriteUInt32(file, 196, checked((uint)bbtOffset));
            file[461] = cryptMethod;
        } else {
            WriteUInt64(file, 216, 0x200);
            WriteUInt64(file, 224, nbtOffset);
            WriteUInt64(file, 232, 0x204);
            WriteUInt64(file, 240, bbtOffset);
            file[512] = 0x80;
            file[513] = cryptMethod;
        }
    }

    private static void WriteBbt(byte[] file, IReadOnlyList<TestBlock> blocks, bool ansi, bool fourK,
        int bbtOffset) {
        int entrySize = ansi ? 12 : 24;
        for (int index = 0; index < blocks.Count; index++) {
            int offset = bbtOffset + index * entrySize;
            TestBlock block = blocks[index];
            if (ansi) {
                WriteUInt32(file, offset, checked((uint)block.Bid));
                WriteUInt32(file, offset + 4, checked((uint)block.Offset));
                WriteUInt16(file, offset + 8, block.StoredData.Length);
                WriteUInt16(file, offset + 10, 1);
            } else {
                WriteUInt64(file, offset, block.Bid);
                WriteUInt64(file, offset + 8, block.Offset);
                WriteUInt16(file, offset + 16, block.StoredData.Length);
                if (fourK) {
                    WriteUInt16(file, offset + 18, block.Data.Length);
                    WriteUInt16(file, offset + 20, 1);
                } else {
                    WriteUInt16(file, offset + 18, 1);
                }
            }
        }
        WriteBTreePageMetadata(
            file, bbtOffset, blocks.Count, entrySize, 0x80, 0x204, ansi, fourK);
    }

    private static void WriteNbt(byte[] file, IReadOnlyList<TestNode> nodes, bool ansi, bool fourK,
        int nbtOffset) {
        int entrySize = ansi ? 16 : 32;
        TestNode[] ordered = nodes.OrderBy(node => node.Nid).ToArray();
        for (int index = 0; index < ordered.Length; index++) {
            int offset = nbtOffset + index * entrySize;
            TestNode node = ordered[index];
            WriteUInt32(file, offset, node.Nid);
            if (ansi) {
                WriteUInt32(file, offset + 4, checked((uint)node.DataBlock.Bid));
                WriteUInt32(file, offset + 8, checked((uint)node.SubnodeBid));
                WriteUInt32(file, offset + 12, node.ParentNid);
            } else {
                WriteUInt64(file, offset + 8, node.DataBlock.Bid);
                WriteUInt64(file, offset + 16, node.SubnodeBid);
                WriteUInt32(file, offset + 24, node.ParentNid);
            }
        }
        WriteBTreePageMetadata(
            file, nbtOffset, nodes.Count, entrySize, 0x81, 0x200, ansi, fourK);
    }

    private static void WriteBTreePageMetadata(byte[] file, int pageOffset, int count,
        int entrySize, byte type, ulong pageBid, bool ansi, bool fourK) {
        int metadataOffset = ansi ? 496 : fourK ? 4056 : 488;
        int trailerOffset = ansi ? 500 : fourK ? 4072 : 496;
        if (fourK) {
            WriteUInt16(file, pageOffset + metadataOffset, count);
            WriteUInt16(file, pageOffset + metadataOffset + 2, metadataOffset / entrySize);
            file[pageOffset + metadataOffset + 4] = checked((byte)entrySize);
            file[pageOffset + metadataOffset + 5] = 0;
        } else {
            file[pageOffset + metadataOffset] = checked((byte)count);
            file[pageOffset + metadataOffset + 1] = checked((byte)(metadataOffset / entrySize));
            file[pageOffset + metadataOffset + 2] = checked((byte)entrySize);
            file[pageOffset + metadataOffset + 3] = 0;
        }
        file[pageOffset + trailerOffset] = type;
        file[pageOffset + trailerOffset + 1] = type;
        WriteUInt16(file, pageOffset + trailerOffset + 2,
            PstSignature.Compute(pageOffset, pageBid));
        uint crc = PstCrc32.Compute(file, pageOffset, trailerOffset);
        if (ansi) {
            WriteUInt32(file, pageOffset + trailerOffset + 4, checked((uint)pageBid));
            WriteUInt32(file, pageOffset + trailerOffset + 8, crc);
        } else {
            WriteUInt32(file, pageOffset + trailerOffset + 4, crc);
            WriteUInt64(file, pageOffset + trailerOffset + 8, pageBid);
            if (fourK) WriteUInt64(file, pageOffset + trailerOffset + 16, 1);
        }
    }

    private static void WriteBlockTrailer(byte[] file, TestBlock block, byte[] storedData,
        bool ansi, bool fourK, int trailerSize) {
        int trailerOffset = block.Offset + block.AllocationLength - trailerSize;
        WriteUInt16(file, trailerOffset, storedData.Length);
        WriteUInt16(file, trailerOffset + 2,
            PstSignature.Compute(block.Offset, block.Bid));
        WriteUInt32(file, trailerOffset + 4, PstCrc32.Compute(storedData));
        if (ansi) WriteUInt32(file, trailerOffset + 8, checked((uint)block.Bid));
        else WriteUInt64(file, trailerOffset + 8, block.Bid);
        if (fourK) {
            WriteUInt16(file, trailerOffset + 16, 2);
            WriteUInt16(file, trailerOffset + 18, block.Data.Length);
        }
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

    private static byte[] CreateStorePropertyContext(uint passwordChecksum) {
        byte[] displayName = Encoding.Unicode.GetBytes("Test Store\0");
        var allocations = new List<byte[]> {
            new byte[] { 0xB5, 0x02, 0x06, 0x00, 0x40, 0x00, 0x00, 0x00 },
            new byte[16],
            displayName
        };
        byte[] records = allocations[1];
        WriteUInt16(records, 0, 0x3001);
        WriteUInt16(records, 2, 0x001F);
        WriteUInt32(records, 4, 0x60);
        WriteUInt16(records, 8, 0x67FF);
        WriteUInt16(records, 10, 0x0003);
        WriteUInt32(records, 12, passwordChecksum);
        return CreateHeapBlock(allocations);
    }

    private static void AddEmbeddedMessageBlocks(TestNode messageNode, ICollection<TestBlock> blocks, bool ansi) {
        const uint attachmentNid = 0x205;
        const uint embeddedMessageNid = 0x224;
        var attachment = new TestBlock(0x110, CreateAttachmentPropertyContext(embeddedMessageNid));
        var embedded = new TestBlock(0x114, CreatePropertyContext(
            (0x001A, "IPM.Note"),
            (0x0037, "Embedded PST message"),
            (0x1000, "Body from the embedded PST item")));
        var attachmentSubnodes = new TestBlock(0x11A,
            CreateSubnodeBlock(embeddedMessageNid, embedded.Bid, 0, ansi));
        var messageSubnodes = new TestBlock(0x11E,
            CreateSubnodeBlock(attachmentNid, attachment.Bid, attachmentSubnodes.Bid, ansi));
        messageNode.SubnodeBid = messageSubnodes.Bid;
        blocks.Add(attachment);
        blocks.Add(embedded);
        blocks.Add(attachmentSubnodes);
        blocks.Add(messageSubnodes);
    }

    private static void AddByValueAttachmentBlocks(TestNode messageNode,
        ICollection<TestBlock> blocks, byte[] content, bool ansi) {
        const uint attachmentNid = 0x205;
        const uint contentNid = 0x321;
        var attachment = new TestBlock(0x110,
            CreateByValueAttachmentPropertyContext(contentNid, content.LongLength));
        var contentBlock = new TestBlock(0x114, content.ToArray());
        var attachmentSubnodes = new TestBlock(0x11A,
            CreateSubnodeBlock(contentNid, contentBlock.Bid, 0, ansi));
        var messageSubnodes = new TestBlock(0x11E,
            CreateSubnodeBlock(attachmentNid, attachment.Bid, attachmentSubnodes.Bid, ansi));
        messageNode.SubnodeBid = messageSubnodes.Bid;
        blocks.Add(attachment);
        blocks.Add(contentBlock);
        blocks.Add(attachmentSubnodes);
        blocks.Add(messageSubnodes);
    }

    private static byte[] CreateAttachmentPropertyContext(uint embeddedMessageNid) {
        byte[] fileName = Encoding.Unicode.GetBytes("forwarded.msg\0");
        var allocations = new List<byte[]> {
            new byte[] { 0xB5, 0x02, 0x06, 0x00, 0x40, 0x00, 0x00, 0x00 },
            new byte[24],
            fileName
        };
        byte[] records = allocations[1];
        WriteUInt16(records, 0, 0x3701);
        WriteUInt16(records, 2, 0x000D);
        WriteUInt32(records, 4, embeddedMessageNid);
        WriteUInt16(records, 8, 0x3705);
        WriteUInt16(records, 10, 0x0003);
        WriteUInt32(records, 12, 5);
        WriteUInt16(records, 16, 0x3707);
        WriteUInt16(records, 18, 0x001F);
        WriteUInt32(records, 20, 0x60);
        return CreateHeapBlock(allocations);
    }

    private static byte[] CreateByValueAttachmentPropertyContext(uint contentNid, long contentLength) {
        if (contentLength > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(contentLength));
        byte[] fileName = Encoding.Unicode.GetBytes("payload.bin\0");
        var allocations = new List<byte[]> {
            new byte[] { 0xB5, 0x02, 0x06, 0x00, 0x40, 0x00, 0x00, 0x00 },
            new byte[32],
            fileName
        };
        byte[] records = allocations[1];
        WriteUInt16(records, 0, 0x0E20);
        WriteUInt16(records, 2, 0x0003);
        WriteUInt32(records, 4, checked((uint)contentLength));
        WriteUInt16(records, 8, 0x3701);
        WriteUInt16(records, 10, 0x0102);
        WriteUInt32(records, 12, contentNid);
        WriteUInt16(records, 16, 0x3705);
        WriteUInt16(records, 18, 0x0003);
        WriteUInt32(records, 20, 1);
        WriteUInt16(records, 24, 0x3707);
        WriteUInt16(records, 26, 0x001F);
        WriteUInt32(records, 28, 0x60);
        return CreateHeapBlock(allocations);
    }

    private static byte[] CreateHeapBlock(IReadOnlyList<byte[]> allocations) {
        int dataEnd = 12 + allocations.Sum(item => item.Length);
        int mapOffset = (dataEnd + 1) & ~1;
        int mapLength = 4 + (allocations.Count + 1) * 2;
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
        WriteUInt16(block, mapOffset, allocations.Count);
        for (int index = 0; index < boundaries.Count; index++) {
            WriteUInt16(block, mapOffset + 4 + index * 2, boundaries[index]);
        }
        return block;
    }

    private static byte[] CreateSubnodeBlock(uint nid, ulong dataBid, ulong subnodeBid, bool ansi) {
        var block = new byte[ansi ? 16 : 32];
        block[0] = 0x02;
        block[1] = 0x00;
        WriteUInt16(block, 2, 1);
        int offset = ansi ? 4 : 8;
        WriteUInt32(block, offset, nid);
        if (ansi) {
            WriteUInt32(block, offset + 4, checked((uint)dataBid));
            WriteUInt32(block, offset + 8, checked((uint)subnodeBid));
        } else {
            WriteUInt64(block, offset + 8, dataBid);
            WriteUInt64(block, offset + 16, subnodeBid);
        }
        return block;
    }

    private static byte[] Compress(byte[] data) {
        using (var output = new MemoryStream()) {
            using (var deflate = new System.IO.Compression.DeflateStream(
                output, System.IO.Compression.CompressionLevel.Optimal, leaveOpen: true)) {
                deflate.Write(data, 0, data.Length);
            }
            return output.ToArray();
        }
    }

    private static int Align(int value, int alignment) => (value + alignment - 1) & ~(alignment - 1);

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
            DataBlock = new TestBlock(0, data);
        }

        internal uint Nid { get; }
        internal uint ParentNid { get; }
        internal TestBlock DataBlock { get; }
        internal ulong SubnodeBid { get; set; }
    }

    private sealed class TestBlock {
        internal TestBlock(ulong bid, byte[] data) {
            Bid = bid;
            Data = data;
        }

        internal ulong Bid { get; set; }
        internal byte[] Data { get; }
        internal byte[] StoredData { get; set; } = Array.Empty<byte>();
        internal int Offset { get; set; }
        internal int AllocationLength { get; set; }
    }
}
