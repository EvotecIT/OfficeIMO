using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO;

/// <summary>
/// Decodes the Cabinet LZX stream carried by a sequence of CFDATA records.
/// The implementation follows the LZX block, tree, match, and E8-translation
/// structures published in Microsoft Open Specifications MS-PATCH and MS-CAB.
/// </summary>
internal sealed class OfficeLzxDecoder {
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

    private readonly CancellationToken _cancellationToken;
    private readonly int _windowSize;
    private readonly int[] _positionBase;
    private readonly int[] _positionFooterBits;
    private readonly byte[] _mainPathLengths;
    private readonly byte[] _lengthPathLengths = new byte[LengthTreeSymbols];
    private readonly byte[] _history;
    private readonly byte[] _frame;
    private readonly int _expectedOutputBytes;
    private int _frameStart;
    private int _frameIndex;
    private OfficeLzxHuffmanTree _mainTree = OfficeLzxHuffmanTree.Empty;
    private OfficeLzxHuffmanTree _lengthTree = OfficeLzxHuffmanTree.Empty;
    private OfficeLzxHuffmanTree _alignedTree = OfficeLzxHuffmanTree.Empty;
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

    /// <summary>Keeps untranslated history in a circular window, independently of the reusable output frame.</summary>
    internal OfficeLzxDecoder(int outputLength, int windowBits, CancellationToken cancellationToken) {
        if (outputLength < 1) throw new ArgumentOutOfRangeException(nameof(outputLength));
        if (windowBits < 15 || windowBits > 21) {
            throw Error("LZX_WINDOW", "The CAB uses an unsupported LZX window size.");
        }

        _windowSize = 1 << windowBits;
        BuildPositionModel(_windowSize, out _positionBase, out _positionFooterBits);
        _mainPathLengths = new byte[LiteralSymbols + PrimaryLengthSymbols * _positionBase.Length];
        int historySize = 1;
        while (historySize < _windowSize && historySize < outputLength) historySize <<= 1;
        _history = new byte[historySize];
        _frame = new byte[Math.Min(32768, outputLength)];
        _expectedOutputBytes = outputLength;
        _cancellationToken = cancellationToken;
    }

    /// <summary>Requires the final CFDATA frame to finish its LZX block and any byte padding.</summary>
    internal void Complete() {
        _cancellationToken.ThrowIfCancellationRequested();
        if (_outputOffset != _expectedOutputBytes || _blockBytesRemaining != 0) {
            throw Error("LZX_CORRUPT", "The LZX stream does not describe the expected expanded size.");
        }
        if (_uncompressedPaddingPending)
            throw Error("LZX_TRUNCATED", "The final LZX uncompressed-block padding byte is missing.");
    }

    /// <summary>Decodes one complete CFDATA frame. The returned buffer is valid only until the next call.
    /// E8 reversal affects output only; later dictionary matches see untranslated history.</summary>
    internal byte[] DecodeFrame(byte[] data, int offset, int count, int expandedSize) {
        if (expandedSize < 1 || expandedSize > _frame.Length) throw new ArgumentOutOfRangeException(nameof(expandedSize));
        _cancellationToken.ThrowIfCancellationRequested();
        _frameStart = _outputOffset;
        int frameEnd = checked(_outputOffset + expandedSize);
        if (frameEnd > _expectedOutputBytes) throw Error("LZX_CORRUPT", "The LZX frames exceed the declared folder size.");
        var reader = new OfficeLzxBitReader(data, offset, count);

        if (_uncompressedPaddingPending) {
            if (reader.RemainingByteCount < 1) {
                throw Error("LZX_TRUNCATED", "The LZX uncompressed-block padding byte is missing.");
            }
            if (reader.ReadRawByte() != 0) throw Error("LZX_CORRUPT", "The LZX uncompressed-block padding is nonzero.");
            _uncompressedPaddingPending = false;
        }

        while (_outputOffset < frameEnd) {
            _cancellationToken.ThrowIfCancellationRequested();
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
            throw Error("LZX_CORRUPT", "An LZX token crosses a 32-KB LZX frame boundary.");
        }
        reader.ValidateFrameEnd();
        if (_translationFileSize > 0 && _frameIndex < MaximumTranslatedFrames)
            ReverseE8Translation(_frame, _frameStart, expandedSize, _translationFileSize);
        _frameIndex++;
        return _frame;
    }

    private void ReadBlockHeader(OfficeLzxBitReader reader) {
        if (!_streamHeaderRead) {
            _streamHeaderRead = true;
            if (reader.ReadBits(1) != 0) {
                uint high = reader.ReadBits(16);
                uint low = reader.ReadBits(16);
                uint fileSize = (high << 16) | low;
                if (fileSize > int.MaxValue) {
                    throw Error("LZX_CORRUPT", "The LZX E8 translation size is outside the supported range.");
                }
                _translationFileSize = (int)fileSize;
            }
        }

        _blockType = (int)reader.ReadBits(3);
        _blockBytesRemaining = checked((int)((reader.ReadBits(16) << 8) | reader.ReadBits(8)));
        if (_blockBytesRemaining == 0 || (_blockType != VerbatimBlock && _blockType != AlignedOffsetBlock && _blockType != UncompressedBlock)) {
            throw Error("LZX_CORRUPT", "The LZX block header contains an invalid type or size.");
        }

        if (_blockType == UncompressedBlock) {
            _uncompressedBlockNeedsPadding = (_blockBytesRemaining & 1) != 0;
            reader.AlignUncompressedHeader();
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
            _alignedTree = OfficeLzxHuffmanTree.Create(alignedLengths, false, "aligned-offset");
        }

        ReadPathLengths(reader, _mainPathLengths, 0, LiteralSymbols);
        ReadPathLengths(reader, _mainPathLengths, LiteralSymbols, _mainPathLengths.Length);
        _mainTree = OfficeLzxHuffmanTree.Create(_mainPathLengths, false, "main");

        ReadPathLengths(reader, _lengthPathLengths, 0, _lengthPathLengths.Length);
        _lengthTree = OfficeLzxHuffmanTree.Create(_lengthPathLengths, true, "length");
    }

    private void DecodeUncompressedSlice(OfficeLzxBitReader reader, int frameEnd) {
        int availableOutput = frameEnd - _outputOffset;
        int copyLength = Math.Min(_blockBytesRemaining, availableOutput);
        int frameOffset = _outputOffset - _frameStart;
        reader.CopyRawBytes(_frame, frameOffset, copyLength);
        int historyOffset = _outputOffset & (_history.Length - 1);
        int first = Math.Min(copyLength, _history.Length - historyOffset);
        Buffer.BlockCopy(_frame, frameOffset, _history, historyOffset, first);
        if (copyLength > first) Buffer.BlockCopy(_frame, frameOffset + first, _history, 0, copyLength - first);
        _outputOffset += copyLength;
        _blockBytesRemaining -= copyLength;
    }

    private void DecodeCompressedSlice(OfficeLzxBitReader reader, int frameEnd) {
        while (_blockBytesRemaining > 0 && _outputOffset < frameEnd) {
            if ((_outputOffset & 1023) == 0) _cancellationToken.ThrowIfCancellationRequested();
            int symbol = _mainTree.Decode(reader);
            if (symbol < LiteralSymbols) {
                WriteByte(checked((byte)symbol));
                _blockBytesRemaining--;
                continue;
            }

            int matchHeader = symbol - LiteralSymbols;
            int positionSlot = matchHeader >> 3;
            int lengthHeader = matchHeader & 7;
            if (positionSlot < 0 || positionSlot >= _positionBase.Length) {
                throw Error("LZX_CORRUPT", "The LZX match references an invalid position slot.");
            }

            int matchLength = lengthHeader + 2;
            if (lengthHeader == 7) matchLength += _lengthTree.Decode(reader);
            uint matchOffset = DecodeMatchOffset(reader, positionSlot);

            if (matchLength > _blockBytesRemaining || matchLength > frameEnd - _outputOffset) {
                throw Error("LZX_CORRUPT", "An LZX match crosses a block or LZX frame boundary.");
            }
            if (matchOffset == 0 || matchOffset > _windowSize || matchOffset > _outputOffset) {
                throw Error("LZX_CORRUPT", "An LZX match references bytes outside the available window.");
            }

            int source = _outputOffset - checked((int)matchOffset);
            for (int index = 0; index < matchLength; index++) {
                WriteByte(_history[(source + index) & (_history.Length - 1)]);
            }
            _blockBytesRemaining -= matchLength;
        }
    }

    private void WriteByte(byte value) {
        _frame[_outputOffset - _frameStart] = value;
        _history[_outputOffset & (_history.Length - 1)] = value;
        _outputOffset++;
    }

    private uint DecodeMatchOffset(OfficeLzxBitReader reader, int positionSlot) {
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
            throw Error("LZX_CORRUPT", "The LZX match contains an invalid formatted offset.");
        }
        uint matchOffset = formattedOffset - 2;
        ValidateRepeatedOffset(matchOffset);
        _recentOffset2 = _recentOffset1;
        _recentOffset1 = _recentOffset0;
        _recentOffset0 = matchOffset;
        return matchOffset;
    }

    private static void ReadPathLengths(OfficeLzxBitReader reader, byte[] lengths, int start, int end) {
        var pretreeLengths = new byte[PretreeSymbols];
        for (int index = 0; index < pretreeLengths.Length; index++) {
            pretreeLengths[index] = checked((byte)reader.ReadBits(4));
        }
        OfficeLzxHuffmanTree pretree = OfficeLzxHuffmanTree.Create(pretreeLengths, false, "pretree");

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
                    throw Error("LZX_CORRUPT", "The LZX path-length repeat contains an invalid delta.");
                }
                value = checked((byte)((lengths[cursor] - delta + 17) % 17));
            } else {
                throw Error("LZX_CORRUPT", "The LZX pretree contains an invalid symbol.");
            }

            if (repeat > end - cursor) {
                throw Error("LZX_CORRUPT", "The LZX path-length repeat exceeds its target tree.");
            }
            for (int index = 0; index < repeat; index++) lengths[cursor++] = value;
        }
    }

    private void ConsumeUncompressedPadding(OfficeLzxBitReader reader) {
        if (!_uncompressedBlockNeedsPadding) return;
        _uncompressedBlockNeedsPadding = false;
        if (reader.RemainingByteCount > 0) {
            if (reader.ReadRawByte() != 0) throw Error("LZX_CORRUPT", "The LZX uncompressed-block padding is nonzero.");
        } else {
            _uncompressedPaddingPending = true;
        }
    }

    private void ValidateRepeatedOffset(uint offset) {
        if (offset == 0 || offset > _windowSize) {
            throw Error("LZX_CORRUPT", "The LZX repeated-offset state is outside the configured window.");
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
        int scanEnd = frameLength - 10;
        for (int cursor = 0; cursor < scanEnd; cursor++) {
            if (data[cursor] != 0xE8) continue;

            int value = ReadInt32(data, cursor + 1);
            int currentPointer = frameOffset + cursor;
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

    private static OfficeLzxException Error(string code, string message) => new OfficeLzxException(code, message);
}
