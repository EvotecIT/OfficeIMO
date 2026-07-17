namespace OfficeIMO.OneNote;

/// <summary>
/// Decodes the Cabinet LZX stream carried by a sequence of CFDATA records.
/// The implementation follows the LZX block, tree, match, and E8-translation
/// structures published in Microsoft Open Specifications MS-PATCH and MS-CAB.
/// </summary>
internal sealed class OneNoteLzxDecoder {
    private const int LiteralSymbols = 256;
    private const int PrimaryLengthSymbols = 8;
    private const int LengthTreeSymbols = 249;
    private const int AlignedTreeSymbols = 8;
    private const int PretreeSymbols = 20;
    private const int MaximumPathLength = 16;
    private const int VerbatimBlock = 1;
    private const int AlignedOffsetBlock = 2;
    private const int UncompressedBlock = 3;
    private const int MaximumTranslatedFrames = 32768;

    private readonly int _windowSize;
    private readonly int[] _positionBase;
    private readonly int[] _positionFooterBits;
    private readonly byte[] _mainPathLengths;
    private readonly byte[] _lengthPathLengths = new byte[LengthTreeSymbols];
    private readonly byte[] _decoded;
    private OneNoteLzxHuffmanTree _mainTree = OneNoteLzxHuffmanTree.Empty;
    private OneNoteLzxHuffmanTree _lengthTree = OneNoteLzxHuffmanTree.Empty;
    private OneNoteLzxHuffmanTree _alignedTree = OneNoteLzxHuffmanTree.Empty;
    private uint _recentOffset0 = 1;
    private uint _recentOffset1 = 1;
    private uint _recentOffset2 = 1;
    private int _outputOffset;
    private int _blockType;
    private int _blockBytesRemaining;
    private bool _streamHeaderRead;
    private int _translationFileSize;
    private bool _uncompressedBlockNeedsPadding;
    private bool _uncompressedPaddingPending;

    private OneNoteLzxDecoder(int outputLength, int windowBits) {
        if (windowBits < 15 || windowBits > 21) {
            throw Error("ONENOTE_CAB_LZX_WINDOW", "The CAB uses an unsupported LZX window size.");
        }

        _windowSize = 1 << windowBits;
        BuildPositionModel(_windowSize, out _positionBase, out _positionFooterBits);
        _mainPathLengths = new byte[LiteralSymbols + PrimaryLengthSymbols * _positionBase.Length];
        _decoded = new byte[outputLength];
    }

    internal static byte[] Decompress(
        IReadOnlyList<byte[]> chunks,
        IReadOnlyList<int> sizes,
        int windowBits,
        long maxOutputBytes) {
        if (chunks == null) throw new ArgumentNullException(nameof(chunks));
        if (sizes == null) throw new ArgumentNullException(nameof(sizes));
        if (chunks.Count != sizes.Count) {
            throw Error("ONENOTE_CAB_LZX_BLOCKS", "CAB LZX block metadata is inconsistent.");
        }
        if (maxOutputBytes < 1) throw new ArgumentOutOfRangeException(nameof(maxOutputBytes));

        long outputLength = 0;
        for (int index = 0; index < sizes.Count; index++) {
            if (chunks[index] == null || sizes[index] < 0 || outputLength > maxOutputBytes - sizes[index]) {
                throw Error("ONENOTE_CAB_EXPANDED_LIMIT", "The expanded CAB folder exceeds the configured size limit.");
            }
            outputLength += sizes[index];
        }
        if (outputLength > int.MaxValue) {
            throw Error("ONENOTE_CAB_EXPANDED_LIMIT", "The expanded CAB folder is too large to materialize.");
        }

        var decoder = new OneNoteLzxDecoder((int)outputLength, windowBits);
        return decoder.Decode(chunks, sizes);
    }

    private byte[] Decode(IReadOnlyList<byte[]> chunks, IReadOnlyList<int> sizes) {
        var frameStarts = new int[sizes.Count];
        for (int frame = 0; frame < chunks.Count; frame++) {
            frameStarts[frame] = _outputOffset;
            DecodeFrame(chunks[frame], sizes[frame]);
        }

        if (_outputOffset != _decoded.Length || _blockBytesRemaining != 0) {
            throw Error("ONENOTE_CAB_LZX_CORRUPT", "The LZX stream does not describe the expected expanded size.");
        }

        byte[] result = (byte[])_decoded.Clone();
        if (_translationFileSize > 0) {
            for (int frame = 0; frame < sizes.Count && frame < MaximumTranslatedFrames; frame++) {
                ReverseE8Translation(result, frameStarts[frame], sizes[frame], _translationFileSize);
            }
        }
        return result;
    }

    private void DecodeFrame(byte[] chunk, int expandedSize) {
        int frameEnd = checked(_outputOffset + expandedSize);
        var reader = new OneNoteLzxBitReader(chunk);

        if (_uncompressedPaddingPending) {
            if (reader.RemainingByteCount < 1) {
                throw Error("ONENOTE_CAB_LZX_TRUNCATED", "The LZX uncompressed-block padding byte is missing.");
            }
            reader.ReadRawByte();
            _uncompressedPaddingPending = false;
        }

        while (_outputOffset < frameEnd) {
            if (_blockBytesRemaining == 0) ReadBlockHeader(reader);

            if (_blockType == UncompressedBlock) {
                DecodeUncompressedSlice(reader, frameEnd);
            } else {
                DecodeCompressedSlice(reader, frameEnd);
            }

            if (_blockBytesRemaining == 0 && _blockType == UncompressedBlock) {
                ConsumeUncompressedPadding(reader);
            }
        }

        if (_outputOffset != frameEnd) {
            throw Error("ONENOTE_CAB_LZX_CORRUPT", "An LZX token crosses a 32-KB CFDATA output boundary.");
        }
    }

    private void ReadBlockHeader(OneNoteLzxBitReader reader) {
        if (!_streamHeaderRead) {
            _streamHeaderRead = true;
            if (reader.ReadBits(1) != 0) {
                uint high = reader.ReadBits(16);
                uint low = reader.ReadBits(16);
                uint fileSize = (high << 16) | low;
                if (fileSize > int.MaxValue) {
                    throw Error("ONENOTE_CAB_LZX_CORRUPT", "The LZX E8 translation size is outside the supported range.");
                }
                _translationFileSize = (int)fileSize;
            }
        }

        _blockType = (int)reader.ReadBits(3);
        _blockBytesRemaining = checked((int)((reader.ReadBits(16) << 8) | reader.ReadBits(8)));
        if (_blockBytesRemaining == 0 || (_blockType != VerbatimBlock && _blockType != AlignedOffsetBlock && _blockType != UncompressedBlock)) {
            throw Error("ONENOTE_CAB_LZX_CORRUPT", "The LZX block header contains an invalid type or size.");
        }

        if (_blockType == UncompressedBlock) {
            _uncompressedBlockNeedsPadding = (_blockBytesRemaining & 1) != 0;
            reader.AlignToWord();
            _recentOffset0 = reader.ReadRawUInt32();
            _recentOffset1 = reader.ReadRawUInt32();
            _recentOffset2 = reader.ReadRawUInt32();
            ValidateRepeatedOffset(_recentOffset0);
            ValidateRepeatedOffset(_recentOffset1);
            ValidateRepeatedOffset(_recentOffset2);
            return;
        }

        if (_blockType == AlignedOffsetBlock) {
            var alignedLengths = new byte[AlignedTreeSymbols];
            for (int index = 0; index < alignedLengths.Length; index++) {
                alignedLengths[index] = checked((byte)reader.ReadBits(3));
            }
            _alignedTree = OneNoteLzxHuffmanTree.Create(alignedLengths, false, "aligned-offset");
        }

        ReadPathLengths(reader, _mainPathLengths, 0, LiteralSymbols);
        ReadPathLengths(reader, _mainPathLengths, LiteralSymbols, _mainPathLengths.Length);
        _mainTree = OneNoteLzxHuffmanTree.Create(_mainPathLengths, false, "main");

        ReadPathLengths(reader, _lengthPathLengths, 0, _lengthPathLengths.Length);
        _lengthTree = OneNoteLzxHuffmanTree.Create(_lengthPathLengths, true, "length");
    }

    private void DecodeUncompressedSlice(OneNoteLzxBitReader reader, int frameEnd) {
        int availableOutput = frameEnd - _outputOffset;
        int copyLength = Math.Min(_blockBytesRemaining, availableOutput);
        reader.CopyRawBytes(_decoded, _outputOffset, copyLength);
        _outputOffset += copyLength;
        _blockBytesRemaining -= copyLength;
    }

    private void DecodeCompressedSlice(OneNoteLzxBitReader reader, int frameEnd) {
        while (_blockBytesRemaining > 0 && _outputOffset < frameEnd) {
            int symbol = _mainTree.Decode(reader);
            if (symbol < LiteralSymbols) {
                _decoded[_outputOffset++] = checked((byte)symbol);
                _blockBytesRemaining--;
                continue;
            }

            int matchHeader = symbol - LiteralSymbols;
            int positionSlot = matchHeader >> 3;
            int lengthHeader = matchHeader & 7;
            if (positionSlot < 0 || positionSlot >= _positionBase.Length) {
                throw Error("ONENOTE_CAB_LZX_CORRUPT", "The LZX match references an invalid position slot.");
            }

            int matchLength = lengthHeader + 2;
            if (lengthHeader == 7) matchLength += _lengthTree.Decode(reader);
            uint matchOffset = DecodeMatchOffset(reader, positionSlot);

            if (matchLength > _blockBytesRemaining || matchLength > frameEnd - _outputOffset) {
                throw Error("ONENOTE_CAB_LZX_CORRUPT", "An LZX match crosses a block or CFDATA output boundary.");
            }
            if (matchOffset == 0 || matchOffset > _windowSize || matchOffset > _outputOffset) {
                throw Error("ONENOTE_CAB_LZX_CORRUPT", "An LZX match references bytes outside the available window.");
            }

            int source = _outputOffset - checked((int)matchOffset);
            for (int index = 0; index < matchLength; index++) {
                _decoded[_outputOffset++] = _decoded[source + index];
            }
            _blockBytesRemaining -= matchLength;
        }
    }

    private uint DecodeMatchOffset(OneNoteLzxBitReader reader, int positionSlot) {
        if (positionSlot == 0) return _recentOffset0;
        if (positionSlot == 1) {
            uint offset = _recentOffset1;
            _recentOffset1 = _recentOffset0;
            _recentOffset0 = offset;
            return offset;
        }
        if (positionSlot == 2) {
            uint offset = _recentOffset2;
            _recentOffset2 = _recentOffset0;
            _recentOffset0 = offset;
            return offset;
        }

        int footerBits = _positionFooterBits[positionSlot];
        uint footer;
        if (_blockType == AlignedOffsetBlock && footerBits >= 3) {
            uint high = footerBits == 3 ? 0 : reader.ReadBits(footerBits - 3) << 3;
            footer = high | checked((uint)_alignedTree.Decode(reader));
        } else {
            footer = reader.ReadBits(footerBits);
        }

        uint formattedOffset = checked((uint)_positionBase[positionSlot] + footer);
        if (formattedOffset < 3) {
            throw Error("ONENOTE_CAB_LZX_CORRUPT", "The LZX match contains an invalid formatted offset.");
        }
        uint matchOffset = formattedOffset - 2;
        ValidateRepeatedOffset(matchOffset);
        _recentOffset2 = _recentOffset1;
        _recentOffset1 = _recentOffset0;
        _recentOffset0 = matchOffset;
        return matchOffset;
    }

    private static void ReadPathLengths(OneNoteLzxBitReader reader, byte[] lengths, int start, int end) {
        var pretreeLengths = new byte[PretreeSymbols];
        for (int index = 0; index < pretreeLengths.Length; index++) {
            pretreeLengths[index] = checked((byte)reader.ReadBits(4));
        }
        OneNoteLzxHuffmanTree pretree = OneNoteLzxHuffmanTree.Create(pretreeLengths, false, "pretree");

        int cursor = start;
        while (cursor < end) {
            int code = pretree.Decode(reader);
            if (code <= MaximumPathLength) {
                lengths[cursor] = checked((byte)((lengths[cursor] - code + 17) % 17));
                cursor++;
                continue;
            }

            int repeat;
            byte value;
            if (code == 17) {
                repeat = checked((int)reader.ReadBits(4) + 4);
                value = 0;
            } else if (code == 18) {
                repeat = checked((int)reader.ReadBits(5) + 20);
                value = 0;
            } else if (code == 19) {
                repeat = checked((int)reader.ReadBits(1) + 4);
                int delta = pretree.Decode(reader);
                if (delta > MaximumPathLength) {
                    throw Error("ONENOTE_CAB_LZX_CORRUPT", "The LZX path-length repeat contains an invalid delta.");
                }
                value = checked((byte)((lengths[cursor] - delta + 17) % 17));
            } else {
                throw Error("ONENOTE_CAB_LZX_CORRUPT", "The LZX pretree contains an invalid symbol.");
            }

            if (repeat > end - cursor) {
                throw Error("ONENOTE_CAB_LZX_CORRUPT", "The LZX path-length repeat exceeds its target tree.");
            }
            for (int index = 0; index < repeat; index++) lengths[cursor++] = value;
        }
    }

    private void ConsumeUncompressedPadding(OneNoteLzxBitReader reader) {
        if (!_uncompressedBlockNeedsPadding) return;
        _uncompressedBlockNeedsPadding = false;
        if (reader.RemainingByteCount > 0) {
            reader.ReadRawByte();
        } else {
            _uncompressedPaddingPending = true;
        }
    }

    private void ValidateRepeatedOffset(uint offset) {
        if (offset == 0 || offset > _windowSize) {
            throw Error("ONENOTE_CAB_LZX_CORRUPT", "The LZX repeated-offset state is outside the configured window.");
        }
    }

    private static void BuildPositionModel(int windowSize, out int[] bases, out int[] footerBits) {
        var baseValues = new List<int>();
        var bitValues = new List<int>();
        int nextBase = 0;
        for (int slot = 0; nextBase < windowSize; slot++) {
            int bits = slot < 4 ? 0 : slot < 36 ? slot / 2 - 1 : 17;
            baseValues.Add(nextBase);
            bitValues.Add(bits);
            nextBase = checked(nextBase + (1 << bits));
        }
        bases = baseValues.ToArray();
        footerBits = bitValues.ToArray();
    }

    private static void ReverseE8Translation(byte[] data, int frameOffset, int frameLength, int fileSize) {
        if (frameOffset >= 0x40000000 || frameLength <= 10) return;
        int scanEnd = checked(frameOffset + frameLength - 10);
        for (int cursor = frameOffset; cursor < scanEnd; cursor++) {
            if (data[cursor] != 0xE8) continue;

            int value = ReadInt32(data, cursor + 1);
            int currentPointer = cursor;
            if (value >= -currentPointer && value < fileSize) {
                int displacement = value >= 0 ? value - currentPointer : value + fileSize;
                WriteInt32(data, cursor + 1, displacement);
            }
            cursor += 4;
        }
    }

    private static int ReadInt32(byte[] data, int offset) => unchecked(
        data[offset] |
        (data[offset + 1] << 8) |
        (data[offset + 2] << 16) |
        (data[offset + 3] << 24));

    private static void WriteInt32(byte[] data, int offset, int value) {
        data[offset] = (byte)value;
        data[offset + 1] = (byte)(value >> 8);
        data[offset + 2] = (byte)(value >> 16);
        data[offset + 3] = (byte)(value >> 24);
    }

    private static OneNoteFormatException Error(string code, string message) => new OneNoteFormatException(code, message);
}
