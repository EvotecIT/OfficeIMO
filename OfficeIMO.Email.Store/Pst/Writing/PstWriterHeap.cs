namespace OfficeIMO.Email.Store;

internal sealed class PstWriterHeap {
    private const int MaximumBlockPayload = 8176;
    private const int AllocationIndexBits = 11;
    private readonly byte _clientSignature;
    private readonly List<HeapBlock> _blocks = new List<HeapBlock>();

    internal PstWriterHeap(byte clientSignature) {
        _clientSignature = clientSignature;
        _blocks.Add(new HeapBlock(0, HeaderSize(0)));
    }

    internal uint Add(byte[] value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        HeapBlock block = _blocks[_blocks.Count - 1];
        if (!block.CanAdd(value.Length)) {
            int index = _blocks.Count;
            if (index > ushort.MaxValue) throw new NotSupportedException("The PST heap exceeds 65,536 blocks.");
            block = new HeapBlock(index, HeaderSize(index));
            _blocks.Add(block);
            if (!block.CanAdd(value.Length)) {
                throw new InvalidOperationException(string.Concat(
                    "A PST heap allocation requires ", value.Length.ToString(CultureInfo.InvariantCulture),
                    " bytes and cannot fit in one heap block."));
            }
        }
        int allocationIndex = block.Add(value);
        if (allocationIndex >= (1 << AllocationIndexBits)) {
            throw new InvalidOperationException("A PST heap block contains too many allocations.");
        }
        return checked(((uint)block.Index << (5 + AllocationIndexBits)) |
            ((uint)(allocationIndex + 1) << 5));
    }

    internal IReadOnlyList<byte[]> Build(uint userRoot) {
        var result = new byte[_blocks.Count][];
        for (int index = 0; index < _blocks.Count; index++) {
            result[index] = BuildBlock(_blocks[index], index + 1 < _blocks.Count, userRoot);
        }
        SetFillLevels(result);
        return result;
    }

    private byte[] BuildBlock(HeapBlock block, bool padToFullBlock, uint userRoot) {
        var allocations = new List<byte[]>(block.Allocations);
        int dataLength = block.HeaderSize + allocations.Sum(item => item.Length);
        if (padToFullBlock) {
            int filler = FindFillerLength(dataLength, allocations.Count);
            if (filler > 0) {
                allocations.Add(new byte[filler]);
                dataLength += filler;
            }
        }
        int mapOffset = (dataLength + 1) & ~1;
        int mapLength = checked(4 + (allocations.Count + 1) * 2);
        int usedLength = checked(mapOffset + mapLength);
        int outputLength = padToFullBlock ? MaximumBlockPayload : usedLength;
        if (usedLength > outputLength || outputLength > MaximumBlockPayload) {
            throw new InvalidOperationException("A PST heap block exceeds its data-block capacity.");
        }
        var bytes = new byte[outputLength];
        PstBinary.WriteUInt16(bytes, 0, mapOffset);
        if (block.Index == 0) {
            bytes[2] = 0xEC;
            bytes[3] = _clientSignature;
            PstBinary.WriteUInt32(bytes, 4, userRoot);
        }
        int cursor = block.HeaderSize;
        PstBinary.WriteUInt16(bytes, mapOffset, allocations.Count);
        PstBinary.WriteUInt16(bytes, mapOffset + 2, 0);
        PstBinary.WriteUInt16(bytes, mapOffset + 4, cursor);
        for (int index = 0; index < allocations.Count; index++) {
            byte[] allocation = allocations[index];
            Buffer.BlockCopy(allocation, 0, bytes, cursor, allocation.Length);
            cursor += allocation.Length;
            PstBinary.WriteUInt16(bytes, mapOffset + 6 + index * 2, cursor);
        }
        return bytes;
    }

    private static int FindFillerLength(int dataLength, int allocationCount) {
        int mapLengthWithFiller = checked(4 + (allocationCount + 2) * 2);
        int maximum = Math.Max(0, MaximumBlockPayload - dataLength - mapLengthWithFiller);
        for (int filler = maximum; filler >= 0; filler--) {
            int used = (((dataLength + filler) + 1) & ~1) + mapLengthWithFiller;
            if (used <= MaximumBlockPayload && MaximumBlockPayload - used <= 63) return filler;
        }
        return 0;
    }

    private static void SetFillLevels(IReadOnlyList<byte[]> blocks) {
        int firstCount = Math.Min(8, blocks.Count);
        for (int index = 0; index < firstCount; index++) {
            SetNibble(blocks[0], 8, index, FillLevel(blocks[index]));
        }
        for (int blockIndex = 8; blockIndex < blocks.Count; blockIndex += 128) {
            int count = Math.Min(128, blocks.Count - blockIndex);
            for (int index = 0; index < count; index++) {
                SetNibble(blocks[blockIndex], 2, index, FillLevel(blocks[blockIndex + index]));
            }
        }
    }

    private static int FillLevel(byte[] block) {
        int free = MaximumBlockPayload - block.Length;
        if (free < 8) return 0xF;
        if (free < 16) return 0xE;
        if (free < 32) return 0xD;
        if (free < 64) return 0xC;
        if (free < 128) return 0xB;
        if (free < 256) return 0xA;
        if (free < 512) return 0x9;
        if (free < 768) return 0x8;
        if (free < 1024) return 0x7;
        if (free < 1280) return 0x6;
        if (free < 1536) return 0x5;
        if (free < 1792) return 0x4;
        if (free < 2048) return 0x3;
        if (free < 2560) return 0x2;
        if (free < 3584) return 0x1;
        return 0x0;
    }

    private static void SetNibble(byte[] bytes, int offset, int index, int value) {
        int byteOffset = offset + index / 2;
        if ((index & 1) == 0) bytes[byteOffset] = checked((byte)((bytes[byteOffset] & 0xF0) | value));
        else bytes[byteOffset] = checked((byte)((bytes[byteOffset] & 0x0F) | (value << 4)));
    }

    private static int HeaderSize(int blockIndex) => blockIndex == 0
        ? 12
        : blockIndex >= 8 && (blockIndex - 8) % 128 == 0 ? 66 : 2;

    private sealed class HeapBlock {
        internal HeapBlock(int index, int headerSize) {
            Index = index;
            HeaderSize = headerSize;
        }
        internal int Index { get; }
        internal int HeaderSize { get; }
        internal List<byte[]> Allocations { get; } = new List<byte[]>();

        internal bool CanAdd(int length) {
            if (Allocations.Count >= (1 << AllocationIndexBits) - 1) return false;
            int dataLength = checked(HeaderSize + Allocations.Sum(item => item.Length) + length);
            int mapOffset = (dataLength + 1) & ~1;
            int mapLength = checked(4 + (Allocations.Count + 2) * 2);
            return mapOffset + mapLength <= MaximumBlockPayload;
        }

        internal int Add(byte[] value) {
            int index = Allocations.Count;
            Allocations.Add(value);
            return index;
        }
    }
}
