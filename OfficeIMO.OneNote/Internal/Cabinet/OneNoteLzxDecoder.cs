namespace OfficeIMO.OneNote;

// CAB LZX decoder ported and adapted from the MIT-licensed implementation in
// deploymenttheory/go-sdk-winmediafoundry/pkg/cab (Copyright 2026 Deployment Theory),
// which in turn is adapted from Microsoft/go-winio. See THIRD-PARTY-NOTICES.md.
internal sealed class OneNoteLzxDecoder {
    private const int MainCodeSplit = 256;
    private const int LengthCodeCount = 249;
    private const int AlignedSymbols = 8;
    private const int PretreeSymbols = 20;
    private const int DecodeBits = 16;
    private const int DecodeSize = 1 << DecodeBits;
    private const int FrameSize = 32768;
    private const int MaxE8Offset = 0x3FFFFFFF;
    private const int MaxPositionSlots = 50;
    private const int MaxMainCode = MainCodeSplit + 8 * MaxPositionSlots;
    private const byte VerbatimBlock = 1;
    private const byte AlignedOffsetBlock = 2;
    private const byte UncompressedBlock = 3;

    private static readonly int[] PositionSlotCounts = { 30, 32, 34, 36, 38, 42, 50 };
    private static readonly byte[] FooterBits = new byte[MaxPositionSlots + 1];
    private static readonly uint[] BasePositions = new uint[MaxPositionSlots + 1];

    private byte[] _input = Array.Empty<byte>();
    private int _inputPosition;
    private byte _bitCount;
    private uint _bits;
    private OneNoteFormatException? _error;
    private bool _skipPaddingByte;
    private readonly uint[] _recentOffsets = { 1, 1, 1 };
    private readonly int _mainElements;
    private bool _headerRead;
    private int _intelFileSize;
    private byte _blockType;
    private uint _blockRemaining;
    private readonly byte[] _mainLengths = new byte[MaxMainCode];
    private readonly byte[] _lengthLengths = new byte[LengthCodeCount];
    private readonly byte[] _alignedLengths = new byte[AlignedSymbols];
    private readonly ushort[] _mainTable = new ushort[DecodeSize];
    private readonly ushort[] _lengthTable = new ushort[DecodeSize];
    private readonly ushort[] _alignedTable = new ushort[DecodeSize];
    private readonly byte[] _output;

    static OneNoteLzxDecoder() {
        uint position = 0;
        for (int index = 0; index <= MaxPositionSlots; index++) {
            byte extraBits;
            if (index < 4) extraBits = 0;
            else if (index < 36) extraBits = (byte)(index / 2 - 1);
            else extraBits = 17;
            FooterBits[index] = extraBits;
            BasePositions[index] = position;
            position += 1U << extraBits;
        }
    }

    private OneNoteLzxDecoder(int outputLength, int windowBits) {
        if (windowBits < 15 || windowBits > 21) {
            throw Error("ONENOTE_CAB_LZX_WINDOW", "The CAB uses an unsupported LZX window size.");
        }
        _mainElements = MainCodeSplit + 8 * PositionSlotCounts[windowBits - 15];
        _output = new byte[outputLength];
    }

    internal static byte[] Decompress(IReadOnlyList<byte[]> chunks, IReadOnlyList<int> sizes, int windowBits, long maxOutputBytes) {
        if (chunks.Count != sizes.Count) throw Error("ONENOTE_CAB_LZX_BLOCKS", "CAB LZX block metadata is inconsistent.");
        long totalLong = 0;
        for (int index = 0; index < sizes.Count; index++) {
            if (sizes[index] < 0 || totalLong > maxOutputBytes - sizes[index]) {
                throw Error("ONENOTE_CAB_EXPANDED_LIMIT", "The expanded CAB folder exceeds the configured size limit.");
            }
            totalLong += sizes[index];
        }
        if (totalLong > int.MaxValue) throw Error("ONENOTE_CAB_EXPANDED_LIMIT", "The expanded CAB folder is too large to materialize.");
        var decoder = new OneNoteLzxDecoder((int)totalLong, windowBits);
        return decoder.Decode(chunks, sizes);
    }

    private byte[] Decode(IReadOnlyList<byte[]> chunks, IReadOnlyList<int> sizes) {
        uint outputPosition = 0;
        for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++) {
            ResetReader(chunks[chunkIndex]);
            if (!_headerRead) {
                if (GetBits(1) != 0) {
                    uint high = GetBits(16);
                    uint low = GetBits(16);
                    _intelFileSize = unchecked((int)((high << 16) | low));
                }
                _headerRead = true;
                ThrowIfFailed();
            }

            uint chunkEnd = checked(outputPosition + (uint)sizes[chunkIndex]);
            while (outputPosition < chunkEnd) {
                if (_blockRemaining == 0) {
                    ReadBlockHeader(out _blockType, out _blockRemaining);
                    if (_blockType != UncompressedBlock) ReadTrees(_blockType == AlignedOffsetBlock);
                }
                uint limit = Math.Min(chunkEnd, checked(outputPosition + _blockRemaining));
                uint produced = _blockType == UncompressedBlock
                    ? CopyUncompressed(outputPosition, limit)
                    : DecodeMatches(outputPosition, limit);
                if (produced == 0) throw Error("ONENOTE_CAB_LZX_CORRUPT", "The CAB LZX stream made no forward progress.");
                outputPosition += produced;
                _blockRemaining -= produced;
            }
        }

        if (_intelFileSize != 0) {
            for (int offset = 0; offset < _output.Length; offset += FrameSize) {
                DecodeE8(_output, offset, Math.Min(FrameSize, _output.Length - offset), offset, _intelFileSize);
            }
        }
        return _output;
    }

    private void ResetReader(byte[] compressed) {
        _input = new byte[compressed.Length + 8];
        Buffer.BlockCopy(compressed, 0, _input, 0, compressed.Length);
        _inputPosition = 0;
        _bitCount = 0;
        _bits = 0;
        _error = null;
    }

    private bool Feed() {
        if (_inputPosition > _input.Length - 2) {
            Fail(Error("ONENOTE_CAB_LZX_TRUNCATED", "The CAB LZX bitstream ended unexpectedly."));
            return false;
        }
        uint word = (uint)(_input[_inputPosition] | (_input[_inputPosition + 1] << 8));
        _bits |= word << (16 - _bitCount);
        _bitCount += 16;
        _inputPosition += 2;
        return true;
    }

    private ushort GetBits(byte count) {
        if (count == 0) return 0;
        if (_bitCount < count && !Feed()) return 0;
        ushort value = (ushort)(_bits >> (32 - count));
        _bits <<= count;
        _bitCount -= count;
        return value;
    }

    private uint GetBitsLong(byte count) {
        if (count <= 16) return GetBits(count);
        return ((uint)GetBits((byte)(count - 16)) << 16) | GetBits(16);
    }

    private static bool BuildTable(byte[] lengths, int lengthOffset, int lengthCount, ushort[] table) {
        Array.Clear(table, 0, table.Length);
        var counts = new int[17];
        for (int index = 0; index < lengthCount; index++) {
            byte length = lengths[lengthOffset + index];
            if (length > 16) return false;
            counts[length]++;
        }
        var positions = new int[18];
        for (int index = 1; index <= 16; index++) positions[index + 1] = positions[index] + (counts[index] << (16 - index));
        for (int symbol = 0; symbol < lengthCount; symbol++) {
            byte length = lengths[lengthOffset + symbol];
            if (length == 0) continue;
            int next = positions[length] + (1 << (16 - length));
            if (next > DecodeSize) return false;
            for (int index = positions[length]; index < next; index++) table[index] = (ushort)symbol;
            positions[length] = next;
        }
        return true;
    }

    private ushort GetCode(ushort[] table, byte[] lengths, int lengthOffset, int lengthCount) {
        if (_bitCount < 16) Feed();
        int symbol = table[_bits >> 16];
        if (symbol >= lengthCount) {
            Fail(Error("ONENOTE_CAB_LZX_CORRUPT", "The CAB LZX Huffman symbol is out of range."));
            return 0;
        }
        byte count = lengths[lengthOffset + symbol];
        if (count == 0 || _bitCount < count) {
            Fail(Error("ONENOTE_CAB_LZX_CORRUPT", "The CAB LZX Huffman code is invalid."));
            return 0;
        }
        _bits <<= count;
        _bitCount -= count;
        return (ushort)symbol;
    }

    private void ReadTree(byte[] lengths, int offset, int count) {
        var pretreeLengths = new byte[PretreeSymbols];
        for (int index = 0; index < pretreeLengths.Length; index++) pretreeLengths[index] = (byte)GetBits(4);
        ThrowIfFailed();
        var pretree = new ushort[DecodeSize];
        if (!BuildTable(pretreeLengths, 0, pretreeLengths.Length, pretree)) throw Error("ONENOTE_CAB_LZX_CORRUPT", "The CAB LZX pretree is over-subscribed.");

        int position = 0;
        while (position < count) {
            byte code = (byte)GetCode(pretree, pretreeLengths, 0, pretreeLengths.Length);
            ThrowIfFailed();
            if (code <= 16) {
                lengths[offset + position] = (byte)((lengths[offset + position] + 17 - code) % 17);
                position++;
            } else if (code == 17 || code == 18) {
                int zeroes = code == 17 ? GetBits(4) + 4 : GetBits(5) + 20;
                if (position > count - zeroes) throw Error("ONENOTE_CAB_LZX_CORRUPT", "The CAB LZX tree zero run exceeds its destination.");
                Array.Clear(lengths, offset + position, zeroes);
                position += zeroes;
            } else if (code == 19) {
                int same = GetBits(1) + 4;
                if (position > count - same) throw Error("ONENOTE_CAB_LZX_CORRUPT", "The CAB LZX tree repeat exceeds its destination.");
                code = (byte)GetCode(pretree, pretreeLengths, 0, pretreeLengths.Length);
                if (code > 16) throw Error("ONENOTE_CAB_LZX_CORRUPT", "The CAB LZX tree repeat code is invalid.");
                byte length = (byte)((lengths[offset + position] + 17 - code) % 17);
                for (int index = 0; index < same; index++) lengths[offset + position + index] = length;
                position += same;
            } else {
                throw Error("ONENOTE_CAB_LZX_CORRUPT", "The CAB LZX tree code is invalid.");
            }
        }
        ThrowIfFailed();
    }

    private void ReadBlockHeader(out byte blockType, out uint blockSize) {
        if (_skipPaddingByte) {
            EnsureInput(1);
            _inputPosition++;
            _skipPaddingByte = false;
        }
        blockType = (byte)GetBits(3);
        uint high = GetBits(16);
        uint low = GetBits(8);
        blockSize = (high << 8) | low;
        ThrowIfFailed();
        if (blockSize == 0) throw Error("ONENOTE_CAB_LZX_CORRUPT", "The CAB LZX block has zero length.");

        if (blockType == UncompressedBlock) {
            byte discard = _bitCount == 0 ? (byte)16 : _bitCount;
            GetBits(discard);
            _bits = 0;
            EnsureInput(12);
            _recentOffsets[0] = ReadUInt32(_input, _inputPosition);
            _recentOffsets[1] = ReadUInt32(_input, _inputPosition + 4);
            _recentOffsets[2] = ReadUInt32(_input, _inputPosition + 8);
            _inputPosition += 12;
            _skipPaddingByte = (blockSize & 1U) != 0;
        } else if (blockType != VerbatimBlock && blockType != AlignedOffsetBlock) {
            throw Error("ONENOTE_CAB_LZX_CORRUPT", "The CAB LZX block type is unsupported.");
        }
    }

    private void ReadTrees(bool aligned) {
        if (aligned) {
            for (int index = 0; index < _alignedLengths.Length; index++) _alignedLengths[index] = (byte)GetBits(3);
            if (!BuildTable(_alignedLengths, 0, _alignedLengths.Length, _alignedTable)) throw Error("ONENOTE_CAB_LZX_CORRUPT", "The CAB LZX aligned tree is invalid.");
        }
        ReadTree(_mainLengths, 0, MainCodeSplit);
        ReadTree(_mainLengths, MainCodeSplit, _mainElements - MainCodeSplit);
        if (!BuildTable(_mainLengths, 0, _mainElements, _mainTable)) throw Error("ONENOTE_CAB_LZX_CORRUPT", "The CAB LZX main tree is invalid.");
        ReadTree(_lengthLengths, 0, _lengthLengths.Length);
        if (!BuildTable(_lengthLengths, 0, _lengthLengths.Length, _lengthTable)) throw Error("ONENOTE_CAB_LZX_CORRUPT", "The CAB LZX length tree is invalid.");
        ThrowIfFailed();
    }

    private uint DecodeMatches(uint start, uint end) {
        bool aligned = _blockType == AlignedOffsetBlock;
        uint outputPosition = start;
        while (outputPosition < end) {
            uint main = GetCode(_mainTable, _mainLengths, 0, _mainElements);
            ThrowIfFailed();
            if (main < 256) {
                _output[outputPosition++] = (byte)main;
                continue;
            }

            uint matchLength = (main - 256) % 8;
            uint slot = (main - 256) / 8;
            if (matchLength == 7) matchLength += GetCode(_lengthTable, _lengthLengths, 0, _lengthLengths.Length);
            matchLength += 2;
            if (slot >= FooterBits.Length) throw Error("ONENOTE_CAB_LZX_CORRUPT", "The CAB LZX match slot is out of range.");

            uint matchOffset;
            if (slot < 3) {
                matchOffset = _recentOffsets[slot];
                _recentOffsets[slot] = _recentOffsets[0];
                _recentOffsets[0] = matchOffset;
            } else {
                byte offsetBits = FooterBits[slot];
                uint verbatimBits = 0;
                uint alignedBits = 0;
                if (aligned && offsetBits >= 3) {
                    verbatimBits = GetBitsLong((byte)(offsetBits - 3)) * 8;
                    alignedBits = GetCode(_alignedTable, _alignedLengths, 0, _alignedLengths.Length);
                } else if (offsetBits > 0) {
                    verbatimBits = GetBitsLong(offsetBits);
                }
                matchOffset = BasePositions[slot] + verbatimBits + alignedBits - 2;
                _recentOffsets[2] = _recentOffsets[1];
                _recentOffsets[1] = _recentOffsets[0];
                _recentOffsets[0] = matchOffset;
            }

            if (matchOffset == 0 || matchOffset > outputPosition || matchLength > end - outputPosition) {
                throw Error("ONENOTE_CAB_LZX_CORRUPT", "The CAB LZX match exceeds the decoded window.");
            }
            uint copyEnd = outputPosition + matchLength;
            while (outputPosition < copyEnd) {
                _output[outputPosition] = _output[outputPosition - matchOffset];
                outputPosition++;
            }
        }
        return outputPosition - start;
    }

    private uint CopyUncompressed(uint start, uint end) {
        int count = checked((int)(end - start));
        EnsureInput(count);
        Buffer.BlockCopy(_input, _inputPosition, _output, checked((int)start), count);
        _inputPosition += count;
        return (uint)count;
    }

    private static void DecodeE8(byte[] data, int dataOffset, int length, int absoluteOffset, int fileSize) {
        if (fileSize == 0 || absoluteOffset > MaxE8Offset || length < 10) return;
        for (int index = 0; index < length - 10; index++) {
            int position = dataOffset + index;
            if (data[position] != 0xE8) continue;
            int currentPointer = absoluteOffset + index;
            int absolute = unchecked((int)ReadUInt32(data, position + 1));
            if (absolute >= -currentPointer && absolute < fileSize) {
                int relative = absolute >= 0 ? absolute - currentPointer : absolute + fileSize;
                WriteUInt32(data, position + 1, unchecked((uint)relative));
            }
            index += 4;
        }
    }

    private void EnsureInput(int count) {
        if (count < 0 || _inputPosition > _input.Length - count) {
            throw Error("ONENOTE_CAB_LZX_TRUNCATED", "The CAB LZX stream ended unexpectedly.");
        }
    }

    private void Fail(OneNoteFormatException error) {
        if (_error == null) _error = error;
    }

    private void ThrowIfFailed() {
        if (_error != null) throw _error;
    }

    private static uint ReadUInt32(byte[] data, int offset) {
        return (uint)(data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24));
    }

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)value;
        data[offset + 1] = (byte)(value >> 8);
        data[offset + 2] = (byte)(value >> 16);
        data[offset + 3] = (byte)(value >> 24);
    }

    private static OneNoteFormatException Error(string code, string message) => new OneNoteFormatException(code, message);
}
