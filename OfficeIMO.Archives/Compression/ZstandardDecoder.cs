namespace OfficeIMO.Archives.Compression;

/// <summary>Decodes Zstandard frames (RFC 8878) from a stream into a destination stream.</summary>
/// <remarks>Memory is the declared window plus the dictionary and two maximum-size blocks, grown only as output is
/// produced. Every structure is validated before it is used and errors name the input offset.</remarks>
internal sealed class ZstandardDecoder {
    internal const string FormatName = "Zstandard";
    internal const uint FrameMagic = 0xFD2FB528;
    internal const uint SkippableMagicMask = 0xFFFFFFF0;
    internal const uint SkippableMagic = 0x184D2A50;
    private const int MaxBlockSize = 128 * 1024;
    private const int LiteralLengthMaxSymbol = 35;
    private const int MatchLengthMaxSymbol = 52;
    private const int OffsetMaxSymbol = 31;

    private static readonly int[] LiteralLengthBaselines = {
        0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15,
        16, 18, 20, 22, 24, 28, 32, 40, 48, 64, 128, 256, 512, 1024, 2048, 4096,
        8192, 16384, 32768, 65536
    };

    private static readonly byte[] LiteralLengthExtraBits = {
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        1, 1, 1, 1, 2, 2, 3, 3, 4, 6, 7, 8, 9, 10, 11, 12,
        13, 14, 15, 16
    };

    private static readonly int[] MatchLengthBaselines = {
        3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18,
        19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34,
        35, 37, 39, 41, 43, 47, 51, 59, 67, 83, 99, 131, 259, 515, 1027, 2051,
        4099, 8195, 16387, 32771, 65539
    };

    private static readonly byte[] MatchLengthExtraBits = {
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5, 7, 8, 9, 10, 11,
        12, 13, 14, 15, 16
    };

    private static readonly ZstandardFseTable DefaultLiteralLengths = ZstandardFseTable.Build(new short[] {
        4, 3, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 3, 2, 1, 1, 1, 1, 1,
        -1, -1, -1, -1
    }, 36, 6, 0, "predefined literal lengths table");

    private static readonly ZstandardFseTable DefaultMatchLengths = ZstandardFseTable.Build(new short[] {
        1, 4, 3, 2, 2, 2, 2, 2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1,
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1, -1, -1, -1, -1, -1, -1
    }, 53, 6, 0, "predefined match lengths table");

    private static readonly ZstandardFseTable DefaultOffsets = ZstandardFseTable.Build(new short[] {
        1, 1, 1, 1, 1, 1, 2, 2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1, -1, -1, -1, -1
    }, 29, 5, 0, "predefined offsets table");

    private readonly Stream _source;
    private readonly Stream _destination;
    private readonly ZstandardDecoderSettings _settings;
    private readonly byte[] _block = new byte[MaxBlockSize];
    private readonly byte[] _literals = new byte[MaxBlockSize];
    private readonly byte[] _small = new byte[8];
    private readonly int[] _repeats = new int[3];
    private readonly ZstandardBackwardBits _bits = new();
    private readonly XxHash64 _checksum = new();
    private byte[] _history = Array.Empty<byte>();
    private long _offset;
    private long _total;

    // Frame state.
    private ZstandardHuffmanTable? _huffman;
    private ZstandardFseTable? _literalLengths;
    private ZstandardFseTable? _offsets;
    private ZstandardFseTable? _matchLengths;
    private long _window;
    private long _capacity;
    private int _blockMaximum;
    private int _dictionaryLength;
    private int _position;
    private long _frameProduced;
    private long _contentSize;
    private long _blockOffset;

    private ZstandardDecoder(Stream source, Stream destination, ZstandardDecoderSettings settings) {
        _source = source;
        _destination = destination;
        _settings = settings;
    }

    internal static ZstandardDecodeResult Decode(Stream source, Stream destination, ZstandardDecoderSettings settings) =>
        new ZstandardDecoder(source, destination, settings).DecodeAll();

    internal static ArchiveCorruptionException Corrupt(long offset, string expected) => new(FormatName, offset, expected);

    internal static int HighBit(uint value) {
        int bit = 0;
        while ((value >>= 1) != 0) bit++;
        return bit;
    }

    private ZstandardDecodeResult DecodeAll() {
        int frames = 0, skippable = 0;
        while (true) {
            _settings.CancellationToken.ThrowIfCancellationRequested();
            long frameOffset = _offset;
            int read = ReadUpTo(_small, 4);
            if (read == 0) {
                if (frames + skippable == 0) throw Corrupt(0, "a Zstandard frame");
                break;
            }
            if (read < 4) throw Corrupt(frameOffset, "a 4-byte frame magic number");
            uint magic = ReadUInt32(_small, 0);
            if ((magic & SkippableMagicMask) == SkippableMagic) {
                SkipFrame(frameOffset);
                skippable++;
                continue;
            }
            if (magic != FrameMagic) throw Corrupt(frameOffset, "the Zstandard or skippable frame magic number");
            _total += DecodeFrame();
            frames++;
        }
        return new ZstandardDecodeResult(frames, skippable, _total);
    }

    private void SkipFrame(long frameOffset) {
        ReadExact(_small, 0, 4, frameOffset + 4, "the skippable frame size");
        long size = ReadUInt32(_small, 0);
        long available = _source.Length - _source.Position;
        if (size > available) throw Corrupt(frameOffset + 8, $"{size} bytes of skippable frame data");
        _source.Position += size;
        _offset += size;
    }

    private long DecodeFrame() {
        long headerOffset = _offset;
        int descriptor = ReadByte("the frame header descriptor");
        int contentSizeFlag = descriptor >> 6;
        bool singleSegment = (descriptor & 0x20) != 0;
        bool hasChecksum = (descriptor & 0x04) != 0;
        int dictionaryFlag = descriptor & 0x03;
        if ((descriptor & 0x08) != 0) throw Corrupt(headerOffset, "the reserved frame header bit to be zero");

        long window = 0;
        if (!singleSegment) {
            long descriptorOffset = _offset;
            int windowDescriptor = ReadByte("the window descriptor");
            int windowLog = 10 + (windowDescriptor >> 3);
            if (windowLog > 41) throw Corrupt(descriptorOffset, "a window of at most 2^41 bytes");
            long windowBase = 1L << windowLog;
            window = windowBase + (windowBase >> 3) * (windowDescriptor & 7);
        }

        uint dictionaryId = 0;
        int dictionaryBytes = dictionaryFlag == 3 ? 4 : dictionaryFlag;
        if (dictionaryBytes > 0) {
            ReadExact(_small, 0, dictionaryBytes, _offset, "the dictionary ID");
            for (int index = 0; index < dictionaryBytes; index++) dictionaryId |= (uint)_small[index] << (8 * index);
        }

        long contentSize = -1;
        int contentSizeBytes = contentSizeFlag switch { 0 => singleSegment ? 1 : 0, 1 => 2, 2 => 4, _ => 8 };
        if (contentSizeBytes > 0) {
            long contentOffset = _offset;
            ReadExact(_small, 0, contentSizeBytes, contentOffset, "the frame content size");
            ulong value = 0;
            for (int index = 0; index < contentSizeBytes; index++) value |= (ulong)_small[index] << (8 * index);
            if (contentSizeBytes == 2) value += 256;
            if (value > long.MaxValue) throw Corrupt(contentOffset, "a frame content size below 2^63");
            contentSize = (long)value;
        }

        long remainingOutput = _settings.MaximumOutput - _total;
        if (contentSize > remainingOutput) throw ArchiveLimits.Exceeded(_settings.OutputLimitName, _settings.MaximumOutput);
        if (singleSegment) {
            // The window is the whole content, which the output limit has just bounded.
            window = contentSize;
        } else if (window > _settings.MaximumWindow) {
            throw ArchiveLimits.Exceeded(nameof(ArchiveReadOptions.MaxWindowBytes), _settings.MaximumWindow);
        }

        ZstandardDictionaryData? dictionary = _settings.Dictionary;
        if (dictionaryId != 0) {
            if (dictionary == null) throw Corrupt(headerOffset, $"dictionary {dictionaryId}, which was not supplied");
            // A raw-content dictionary has ID zero, so it never matches a frame that names a dictionary.
            if (dictionary.Id != dictionaryId)
                throw Corrupt(headerOffset, $"dictionary {dictionaryId}, not the supplied dictionary {dictionary.Id}");
        }
        StartFrame(window, contentSize, dictionary);

        bool last;
        do {
            _settings.CancellationToken.ThrowIfCancellationRequested();
            long blockOffset = _offset;
            ReadExact(_small, 0, 3, blockOffset, "a 3-byte block header");
            int header = _small[0] | _small[1] << 8 | _small[2] << 16;
            last = (header & 1) != 0;
            int type = (header >> 1) & 3;
            int size = header >> 3;
            _blockOffset = blockOffset;
            EnsureRoom();
            int start = _position;
            switch (type) {
                case 0:
                    CheckBlockSize(size, blockOffset);
                    ReadExact(_history, _position, size, blockOffset + 3, "the raw block content");
                    _position += size;
                    break;
                case 1:
                    CheckBlockSize(size, blockOffset);
                    byte value = (byte)ReadByte("the RLE block byte");
                    for (int index = 0; index < size; index++) _history[_position++] = value;
                    break;
                case 2:
                    CheckBlockSize(size, blockOffset);
                    ReadExact(_block, 0, size, blockOffset + 3, "the compressed block content");
                    DecodeCompressedBlock(size, blockOffset + 3);
                    break;
                default:
                    throw Corrupt(blockOffset, "a block type other than the reserved type 3");
            }
            Publish(start, type == 1 ? 1 : size);
        } while (!last);

        if (contentSize >= 0 && _frameProduced != contentSize)
            throw Corrupt(_offset, $"{contentSize} bytes of frame content as declared, not {_frameProduced}");
        if (hasChecksum) {
            long checksumOffset = _offset;
            ReadExact(_small, 0, 4, checksumOffset, "the 4-byte content checksum");
            if (ReadUInt32(_small, 0) != unchecked((uint)_checksum.Complete()))
                throw Corrupt(checksumOffset, "a content checksum that matches the decoded frame");
        }
        return _frameProduced;
    }

    private void StartFrame(long window, long contentSize, ZstandardDictionaryData? dictionary) {
        byte[] content = dictionary?.Content ?? Array.Empty<byte>();
        _dictionaryLength = content.Length;
        _window = window;
        _blockMaximum = (int)Math.Min(window, MaxBlockSize);
        // A frame never needs more history than its window or, when declared, its whole content.
        long history = contentSize >= 0 ? Math.Min(window, contentSize) : window;
        // Spare room of at least half the window means a slide, which copies one window, happens at most once per half
        // window of output, so history copies stay below twice the output.
        _capacity = content.Length + history + Math.Max(2L * MaxBlockSize, (content.Length + history) / 2);
        if (_capacity > int.MaxValue - 64)
            throw ArchiveLimits.Exceeded(nameof(ArchiveReadOptions.MaxWindowBytes), _settings.MaximumWindow);
        int initial = (int)Math.Min(_capacity, content.Length + 2L * MaxBlockSize);
        if (_history.Length < initial) _history = new byte[initial];
        // Every frame starts from a fresh copy of the dictionary, so the copy is decode work like any block.
        if (content.Length > 0) _settings.Budget.Charge(0, content.Length);
        Buffer.BlockCopy(content, 0, _history, 0, content.Length);
        _position = content.Length;
        _frameProduced = 0;
        _contentSize = contentSize;
        _checksum.Reset();
        _huffman = dictionary?.Huffman;
        _literalLengths = dictionary?.LiteralLengths;
        _offsets = dictionary?.Offsets;
        _matchLengths = dictionary?.MatchLengths;
        if (dictionary != null) {
            Array.Copy(dictionary.Repeats, _repeats, 3);
        } else {
            _repeats[0] = 1;
            _repeats[1] = 4;
            _repeats[2] = 8;
        }
    }

    private void CheckBlockSize(int size, long blockOffset) {
        if (size > _blockMaximum) throw Corrupt(blockOffset, $"a block of at most {_blockMaximum} bytes");
    }

    /// <summary>Makes room for one more block, growing the history towards its capacity and then sliding it so the
    /// last window of output stays addressable.</summary>
    private void EnsureRoom() {
        if (_history.Length - _position >= _blockMaximum) return;
        if (_history.Length < _capacity) {
            long grown = Math.Min(_capacity, Math.Max(2L * _history.Length, (long)_position + _blockMaximum));
            Array.Resize(ref _history, (int)grown);
            return;
        }
        int keep = (int)Math.Min(_position, _window);
        _settings.Budget.Charge(0, keep);
        Buffer.BlockCopy(_history, _position - keep, _history, 0, keep);
        _position = keep;
        // Only a frame that outgrows its declared content can run out of history capacity.
        if (_history.Length - _position < _blockMaximum)
            throw Corrupt(_blockOffset, $"no more than the declared {_contentSize} bytes of frame content");
    }

    private void Publish(int start, int compressedBytes) {
        int produced = _position - start;
        long work = Math.Max(4096, (long)compressedBytes + produced);
        if (produced > _settings.MaximumOutput - _total - _frameProduced)
            throw ArchiveLimits.Exceeded(_settings.OutputLimitName, _settings.MaximumOutput);
        if (_contentSize >= 0 && produced > _contentSize - _frameProduced)
            throw Corrupt(_blockOffset, $"no more than the declared {_contentSize} bytes of frame content");
        _settings.Budget.Charge(produced, work);
        _settings.Select?.Invoke(produced);
        if (produced > 0) {
            _destination.Write(_history, start, produced);
            _checksum.Append(_history, start, produced);
        }
        _frameProduced += produced;
    }

    private void DecodeCompressedBlock(int size, long offset) {
        int literalCount = DecodeLiterals(size, offset, out int position);
        int blockStart = _position;
        int limit = blockStart + _blockMaximum;
        if (position >= size) throw Corrupt(offset + position, "the sequences section header");
        int sequenceCount = _block[position++];
        if (sequenceCount >= 128) {
            if (sequenceCount < 255) {
                if (position >= size) throw Corrupt(offset + position, "the sequence count");
                sequenceCount = ((sequenceCount - 128) << 8) + _block[position++];
            } else {
                if (position + 2 > size) throw Corrupt(offset + position, "the sequence count");
                sequenceCount = _block[position] + (_block[position + 1] << 8) + 0x7F00;
                position += 2;
            }
        }
        // Zero can also be written in two bytes (0x80 0x00); either way the block is its literals alone.
        if (sequenceCount == 0) {
            if (position != size) throw Corrupt(offset + position, "the end of a block without sequences");
            if (literalCount > limit - _position) throw Corrupt(offset, $"a block that expands to at most {_blockMaximum} bytes");
            Buffer.BlockCopy(_literals, 0, _history, _position, literalCount);
            _position += literalCount;
            return;
        }
        if (position >= size) throw Corrupt(offset + position, "the symbol compression modes");
        int modes = _block[position++];
        if ((modes & 3) != 0) throw Corrupt(offset + position - 1, "the reserved compression mode bits to be zero");
        _literalLengths = SelectTable(modes >> 6, ref position, size, offset, DefaultLiteralLengths, LiteralLengthMaxSymbol, 9,
            _literalLengths, "literal lengths table");
        _offsets = SelectTable((modes >> 4) & 3, ref position, size, offset, DefaultOffsets, OffsetMaxSymbol, 8,
            _offsets, "offsets table");
        _matchLengths = SelectTable((modes >> 2) & 3, ref position, size, offset, DefaultMatchLengths, MatchLengthMaxSymbol, 9,
            _matchLengths, "match lengths table");
        ExecuteSequences(sequenceCount, position, size, offset, literalCount, limit);
    }

    private ZstandardFseTable SelectTable(int mode, ref int position, int size, long offset, ZstandardFseTable predefined,
        int maxSymbol, int maxAccuracyLog, ZstandardFseTable? previous, string structure) {
        switch (mode) {
            case 0:
                return predefined;
            case 1:
                if (position >= size) throw Corrupt(offset + position, "the RLE symbol of the " + structure);
                int symbol = _block[position];
                if (symbol > maxSymbol) throw Corrupt(offset + position, $"an RLE symbol of at most {maxSymbol} for the {structure}");
                position++;
                return ZstandardFseTable.Rle((byte)symbol);
            case 2:
                position += ZstandardFseTable.Read(_block, position, size, maxSymbol, maxAccuracyLog, offset + position,
                    structure, out ZstandardFseTable table);
                return table;
            default:
                return previous ?? throw Corrupt(offset + position, "a previous " + structure + " to repeat");
        }
    }

    private int DecodeLiterals(int size, long offset, out int position) {
        int first = _block[0];
        int type = first & 3;
        int sizeFormat = (first >> 2) & 3;
        int headerBytes, regenerated, compressed = 0, streams = 1;
        if (type < 2) {
            switch (sizeFormat) {
                case 1:
                    headerBytes = 2;
                    break;
                case 3:
                    headerBytes = 3;
                    break;
                default:
                    headerBytes = 1;
                    break;
            }
            if (headerBytes > size) throw Corrupt(offset, "a complete literals section header");
            regenerated = headerBytes switch {
                1 => first >> 3,
                2 => (first >> 4) + (_block[1] << 4),
                _ => (first >> 4) + (_block[1] << 4) + (_block[2] << 12)
            };
        } else {
            headerBytes = sizeFormat < 2 ? 3 : sizeFormat + 2;
            if (sizeFormat != 0) streams = 4;
            if (headerBytes > size) throw Corrupt(offset, "a complete literals section header");
            long value = 0;
            for (int index = 0; index < headerBytes; index++) value |= (long)_block[index] << (8 * index);
            switch (headerBytes) {
                case 3:
                    regenerated = (int)((value >> 4) & 0x3FF);
                    compressed = (int)((value >> 14) & 0x3FF);
                    break;
                case 4:
                    regenerated = (int)((value >> 4) & 0x3FFF);
                    compressed = (int)((value >> 18) & 0x3FFF);
                    break;
                default:
                    regenerated = (int)((value >> 4) & 0x3FFFF);
                    compressed = (int)((value >> 22) & 0x3FFFF);
                    break;
            }
        }
        if (regenerated > _blockMaximum) throw Corrupt(offset, $"at most {_blockMaximum} literals in a block");
        switch (type) {
            case 0:
                if (regenerated > size - headerBytes) throw Corrupt(offset + headerBytes, $"{regenerated} raw literals");
                Buffer.BlockCopy(_block, headerBytes, _literals, 0, regenerated);
                position = headerBytes + regenerated;
                return regenerated;
            case 1:
                if (headerBytes >= size) throw Corrupt(offset + headerBytes, "the RLE literal byte");
                byte value = _block[headerBytes];
                for (int index = 0; index < regenerated; index++) _literals[index] = value;
                position = headerBytes + 1;
                return regenerated;
        }
        if (compressed > size - headerBytes) throw Corrupt(offset + headerBytes, $"{compressed} bytes of compressed literals");
        int start = headerBytes;
        int end = headerBytes + compressed;
        if (type == 2) {
            start += ZstandardHuffmanTable.Read(_block, start, end, offset + start, _bits, out ZstandardHuffmanTable table);
            _huffman = table;
        } else if (_huffman == null) {
            throw Corrupt(offset, "a previous Huffman table for treeless literals");
        }
        if (streams == 1) {
            DecodeLiteralStream(start, end - start, 0, regenerated, offset + start);
        } else {
            if (end - start < 6) throw Corrupt(offset + start, "the 6-byte literal jump table");
            int size1 = _block[start] | _block[start + 1] << 8;
            int size2 = _block[start + 2] | _block[start + 3] << 8;
            int size3 = _block[start + 4] | _block[start + 5] << 8;
            int size4 = end - start - 6 - size1 - size2 - size3;
            int segment = (regenerated + 3) / 4;
            int lastSegment = regenerated - 3 * segment;
            if (size4 <= 0 || lastSegment < 0) throw Corrupt(offset + start, "a literal jump table within the literals section");
            int stream = start + 6;
            DecodeLiteralStream(stream, size1, 0, segment, offset + stream);
            stream += size1;
            DecodeLiteralStream(stream, size2, segment, segment, offset + stream);
            stream += size2;
            DecodeLiteralStream(stream, size3, 2 * segment, segment, offset + stream);
            stream += size3;
            DecodeLiteralStream(stream, size4, 3 * segment, lastSegment, offset + stream);
        }
        position = end;
        return regenerated;
    }

    private void DecodeLiteralStream(int start, int length, int outputStart, int count, long offset) {
        ZstandardHuffmanTable table = _huffman!;
        _bits.Reset(_block, start, length, offset, "literal stream");
        int maxBits = table.MaxBits;
        byte[] symbols = table.Symbols;
        byte[] lengths = table.Lengths;
        byte[] literals = _literals;
        for (int index = outputStart, end = outputStart + count; index < end; index++) {
            uint code = _bits.Peek(maxBits);
            literals[index] = symbols[code];
            _bits.Skip(lengths[code]);
        }
        if (_bits.Remaining != 0) throw Corrupt(offset, $"a literal stream holding exactly {count} literals");
    }

    private void ExecuteSequences(int count, int position, int size, long offset, int literalCount, int limit) {
        ZstandardFseTable literalLengths = _literalLengths!, offsets = _offsets!, matchLengths = _matchLengths!;
        ZstandardBackwardBits bits = _bits;
        bits.Reset(_block, position, size - position, offset + position, "sequence stream");
        int literalState = (int)bits.Read(literalLengths.AccuracyLog);
        int offsetState = (int)bits.Read(offsets.AccuracyLog);
        int matchState = (int)bits.Read(matchLengths.AccuracyLog);
        byte[] history = _history;
        int blockStart = _position;
        int literal = 0;
        for (int sequence = 0; sequence < count; sequence++) {
            int offsetCode = offsets.Symbols[offsetState];
            int matchCode = matchLengths.Symbols[matchState];
            int literalCode = literalLengths.Symbols[literalState];
            long offsetValue = (1L << offsetCode) + bits.Read(offsetCode);
            int matchLength = MatchLengthBaselines[matchCode] + (int)bits.Read(MatchLengthExtraBits[matchCode]);
            int literalLength = LiteralLengthBaselines[literalCode] + (int)bits.Read(LiteralLengthExtraBits[literalCode]);
            if (sequence != count - 1) {
                literalState = literalLengths.Baselines[literalState] + (int)bits.Read(literalLengths.Bits[literalState]);
                matchState = matchLengths.Baselines[matchState] + (int)bits.Read(matchLengths.Bits[matchState]);
                offsetState = offsets.Baselines[offsetState] + (int)bits.Read(offsets.Bits[offsetState]);
            }
            if (bits.Overflowed) throw Corrupt(offset + position, $"a sequence stream holding {count} sequences");

            long matchOffset = ResolveOffset(offsetValue, literalLength);
            if (literalLength > literalCount - literal)
                throw Corrupt(offset + position, "literal lengths within the block's literals");
            if ((long)literalLength + matchLength > limit - _position)
                throw Corrupt(offset, $"a block that expands to at most {_blockMaximum} bytes");
            Buffer.BlockCopy(_literals, literal, history, _position, literalLength);
            literal += literalLength;
            _position += literalLength;

            long produced = _frameProduced + (_position - blockStart);
            long reach = produced <= _window ? produced + _dictionaryLength : _window;
            if (matchOffset <= 0 || matchOffset > reach) throw Corrupt(offset + position, $"a match offset of at most {reach}");
            int source = _position - (int)matchOffset;
            if (matchOffset >= matchLength) {
                Buffer.BlockCopy(history, source, history, _position, matchLength);
                _position += matchLength;
            } else {
                for (int index = 0; index < matchLength; index++) history[_position++] = history[source++];
            }
        }
        if (bits.Remaining != 0) throw Corrupt(offset + position, $"a sequence stream holding exactly {count} sequences");
        int rest = literalCount - literal;
        if (rest > limit - _position) throw Corrupt(offset, $"a block that expands to at most {_blockMaximum} bytes");
        Buffer.BlockCopy(_literals, literal, history, _position, rest);
        _position += rest;
    }

    private long ResolveOffset(long offsetValue, int literalLength) {
        int[] repeats = _repeats;
        if (offsetValue > 3) {
            long value = offsetValue - 3;
            repeats[2] = repeats[1];
            repeats[1] = repeats[0];
            repeats[0] = (int)Math.Min(value, int.MaxValue);
            return value;
        }
        int index = (int)offsetValue - 1 + (literalLength == 0 ? 1 : 0);
        if (index == 0) return repeats[0];
        int selected = index == 3 ? repeats[0] - 1 : repeats[index];
        if (index != 1) repeats[2] = repeats[1];
        repeats[1] = repeats[0];
        repeats[0] = selected;
        return selected;
    }

    private int ReadByte(string structure) {
        int value = _source.ReadByte();
        if (value < 0) throw Corrupt(_offset, structure);
        _offset++;
        return value;
    }

    private void ReadExact(byte[] buffer, int offset, int count, long inputOffset, string structure) {
        int total = 0;
        while (total < count) {
            int read = _source.Read(buffer, offset + total, count - total);
            if (read <= 0) throw Corrupt(inputOffset, $"{count} bytes of {structure}, found {total}");
            total += read;
        }
        _offset += count;
    }

    private int ReadUpTo(byte[] buffer, int count) {
        int total = 0;
        while (total < count) {
            int read = _source.Read(buffer, total, count - total);
            if (read <= 0) break;
            total += read;
        }
        _offset += total;
        return total;
    }

    private static uint ReadUInt32(byte[] data, int offset) =>
        (uint)(data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16 | data[offset + 3] << 24);
}

internal sealed class ZstandardDecoderSettings {
    internal ZstandardDecoderSettings(long maximumOutput, string outputLimitName, long maximumWindow, ArchiveBlockBudget budget,
        Action<long>? select, ZstandardDictionaryData? dictionary, CancellationToken cancellationToken) {
        MaximumOutput = maximumOutput;
        OutputLimitName = outputLimitName;
        MaximumWindow = maximumWindow;
        Budget = budget;
        Select = select;
        Dictionary = dictionary;
        CancellationToken = cancellationToken;
    }

    internal long MaximumOutput { get; }
    internal string OutputLimitName { get; }
    internal long MaximumWindow { get; }
    internal ArchiveBlockBudget Budget { get; }
    internal Action<long>? Select { get; }
    internal ZstandardDictionaryData? Dictionary { get; }
    internal CancellationToken CancellationToken { get; }
}

internal readonly struct ZstandardDecodeResult {
    internal ZstandardDecodeResult(int frames, int skippableFrames, long length) {
        Frames = frames;
        SkippableFrames = skippableFrames;
        Length = length;
    }

    internal int Frames { get; }
    internal int SkippableFrames { get; }
    internal long Length { get; }
}

/// <summary>A parsed Zstandard dictionary (RFC 8878, section 5): raw content, or content with entropy tables and
/// repeat offsets.</summary>
internal sealed class ZstandardDictionaryData {
    private const uint DictionaryMagic = 0xEC30A437;
    private const string FormatName = "Zstandard dictionary";

    private ZstandardDictionaryData(uint id, byte[] content, int[] repeats, ZstandardHuffmanTable? huffman,
        ZstandardFseTable? literalLengths, ZstandardFseTable? offsets, ZstandardFseTable? matchLengths) {
        Id = id;
        Content = content;
        Repeats = repeats;
        Huffman = huffman;
        LiteralLengths = literalLengths;
        Offsets = offsets;
        MatchLengths = matchLengths;
    }

    internal uint Id { get; }
    internal byte[] Content { get; }
    internal int[] Repeats { get; }
    internal ZstandardHuffmanTable? Huffman { get; }
    internal ZstandardFseTable? LiteralLengths { get; }
    internal ZstandardFseTable? Offsets { get; }
    internal ZstandardFseTable? MatchLengths { get; }
    internal bool HasEntropyTables => Huffman != null;

    internal static ZstandardDictionaryData Parse(byte[] bytes) {
        if (bytes.Length >= 4 && Read32(bytes, 0) == DictionaryMagic && bytes.Length < 8)
            throw new ArchiveCorruptionException(FormatName, 4, "a 4-byte dictionary ID after the dictionary magic");
        if (bytes.Length < 8 || Read32(bytes, 0) != DictionaryMagic) {
            if (bytes.Length == 0) throw new ArchiveCorruptionException(FormatName, 0, "non-empty dictionary content");
            return new ZstandardDictionaryData(0, (byte[])bytes.Clone(), new[] { 1, 4, 8 }, null, null, null, null);
        }
        uint id = Read32(bytes, 4);
        if (id == 0) throw new ArchiveCorruptionException(FormatName, 4, "a non-zero dictionary ID");
        int position = 8;
        try {
            position += ZstandardHuffmanTable.Read(bytes, position, bytes.Length, position, new ZstandardBackwardBits(), out ZstandardHuffmanTable huffman);
            position += ZstandardFseTable.Read(bytes, position, bytes.Length, 31, 8, position, "offsets table", out ZstandardFseTable offsets);
            position += ZstandardFseTable.Read(bytes, position, bytes.Length, 52, 9, position, "match lengths table", out ZstandardFseTable matchLengths);
            position += ZstandardFseTable.Read(bytes, position, bytes.Length, 35, 9, position, "literal lengths table", out ZstandardFseTable literalLengths);
            if (bytes.Length - position < 12) throw new ArchiveCorruptionException(FormatName, position, "three 4-byte repeat offsets");
            var repeats = new int[3];
            int contentLength = bytes.Length - position - 12;
            for (int index = 0; index < 3; index++) {
                uint value = Read32(bytes, position + 4 * index);
                if (value == 0 || value > contentLength)
                    throw new ArchiveCorruptionException(FormatName, position + 4 * index, $"a repeat offset between 1 and the content length {contentLength}");
                repeats[index] = (int)value;
            }
            position += 12;
            var content = new byte[contentLength];
            Buffer.BlockCopy(bytes, position, content, 0, contentLength);
            return new ZstandardDictionaryData(id, content, repeats, huffman, literalLengths, offsets, matchLengths);
        } catch (ArchiveCorruptionException error) when (error.Format != FormatName) {
            throw new ArchiveCorruptionException(FormatName, error.Offset, error.Expected, error);
        }
    }

    private static uint Read32(byte[] data, int offset) =>
        (uint)(data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16 | data[offset + 3] << 24);
}
