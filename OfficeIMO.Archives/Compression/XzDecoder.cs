using System.Security.Cryptography;
using OfficeIMO.Archives.SevenZip;
using OfficeIMO.Drawing;

namespace OfficeIMO.Archives.Compression;

/// <summary>Decodes .xz files (The .xz File Format 1.2.1): streams, blocks with LZMA2 and the delta filter, integrity
/// checks, the index and stream padding.</summary>
internal sealed class XzDecoder {
    internal const string FormatName = "xz";
    private static readonly byte[] HeaderMagic = { 0xFD, 0x37, 0x7A, 0x58, 0x5A, 0x00 };
    private const int FilterDelta = 0x03;
    private const int FilterLzma2 = 0x21;

    private readonly Stream _source;
    private readonly Stream _destination;
    private readonly CompressedStreamSettings _settings;
    private readonly long _start;
    private readonly byte[] _buffer = new byte[1024];
    private long _total;

    private XzDecoder(Stream source, Stream destination, CompressedStreamSettings settings) {
        _source = source;
        _destination = destination;
        _settings = settings;
        _start = source.Position;
    }

    internal static XzDecodeResult Decode(Stream source, Stream destination, CompressedStreamSettings settings) =>
        new XzDecoder(source, destination, settings).DecodeAll();

    private long Offset => _source.Position - _start;

    private static ArchiveCorruptionException Corrupt(long offset, string expected) => new(FormatName, offset, expected);

    private XzDecodeResult DecodeAll() {
        int streams = 0, blocks = 0;
        XzCheck check = XzCheck.None;
        long end = _source.Length;
        if (_source.Position == end) throw Corrupt(0, "the xz stream header");
        while (_source.Position < end) {
            _settings.CancellationToken.ThrowIfCancellationRequested();
            if (streams > 0 && SkipStreamPadding()) break;
            check = DecodeStream(ref blocks);
            streams++;
        }
        return new XzDecodeResult(streams, blocks, check, _total);
    }

    /// <summary>Skips null stream padding, which must come in multiples of four bytes. Returns true at the end of input.</summary>
    private bool SkipStreamPadding() {
        long start = Offset;
        long padding = 0;
        while (true) {
            int value = _source.ReadByte();
            if (value < 0) break;
            if (value != 0) {
                _source.Position--;
                break;
            }
            padding++;
        }
        if (padding % 4 != 0) throw Corrupt(start, "stream padding in multiples of four null bytes");
        return _source.Position == _source.Length;
    }

    private XzCheck DecodeStream(ref int blockTotal) {
        long headerOffset = Offset;
        ReadExact(_buffer, 12, headerOffset, "the 12-byte stream header");
        for (int index = 0; index < HeaderMagic.Length; index++)
            if (_buffer[index] != HeaderMagic[index]) throw Corrupt(headerOffset, "the xz stream header magic bytes");
        if (OfficePngCrc32.Compute(_buffer, 6, 2) != ReadUInt32(_buffer, 8)) throw Corrupt(headerOffset + 8, "a stream flags CRC32 that matches");
        if (_buffer[6] != 0 || (_buffer[7] & 0xF0) != 0) throw Corrupt(headerOffset + 6, "reserved stream flag bits to be zero");
        byte flags0 = _buffer[6], flags1 = _buffer[7];
        var check = (XzCheck)(flags1 & 0x0F);
        if (check is not (XzCheck.None or XzCheck.Crc32 or XzCheck.Crc64 or XzCheck.Sha256))
            throw new NotSupportedException($"xz integrity check type 0x{(int)check:x} is unsupported.");

        var records = new List<(long Unpadded, long Uncompressed)>();
        while (true) {
            _settings.CancellationToken.ThrowIfCancellationRequested();
            long blockOffset = Offset;
            int sizeByte = ReadByte(blockOffset, "a block header or the index");
            if (sizeByte == 0) break;
            records.Add(DecodeBlock(blockOffset, sizeByte, check));
            blockTotal++;
        }
        long indexOffset = Offset - 1;
        long indexSize = ReadIndex(indexOffset, records);

        long footerOffset = Offset;
        ReadExact(_buffer, 12, footerOffset, "the 12-byte stream footer");
        if (_buffer[10] != (byte)'Y' || _buffer[11] != (byte)'Z') throw Corrupt(footerOffset + 10, "the stream footer magic bytes YZ");
        if (OfficePngCrc32.Compute(_buffer, 4, 6) != ReadUInt32(_buffer, 0)) throw Corrupt(footerOffset, "a stream footer CRC32 that matches");
        if (_buffer[8] != flags0 || _buffer[9] != flags1) throw Corrupt(footerOffset + 8, "stream footer flags equal to the header flags");
        long backwardSize = ((long)ReadUInt32(_buffer, 4) + 1) * 4;
        if (backwardSize != indexSize) throw Corrupt(footerOffset + 4, $"a backward size equal to the index size {indexSize}");
        return check;
    }

    private (long Unpadded, long Uncompressed) DecodeBlock(long blockOffset, int sizeByte, XzCheck check) {
        int headerSize = (sizeByte + 1) * 4;
        _buffer[0] = (byte)sizeByte;
        ReadExact(_buffer, 1, headerSize - 1, blockOffset + 1, "the block header");
        if (OfficePngCrc32.Compute(_buffer, 0, headerSize - 4) != ReadUInt32(_buffer, headerSize - 4))
            throw Corrupt(blockOffset + headerSize - 4, "a block header CRC32 that matches");
        int flags = _buffer[1];
        if ((flags & 0x3C) != 0) throw Corrupt(blockOffset + 1, "reserved block flag bits to be zero");
        int filterCount = (flags & 3) + 1;
        int position = 2;
        long declaredCompressed = -1, declaredUncompressed = -1;
        if ((flags & 0x40) != 0) {
            declaredCompressed = ReadVarint(_buffer, ref position, headerSize - 4, blockOffset, "the block compressed size");
            if (declaredCompressed == 0) throw Corrupt(blockOffset + 2, "a non-zero block compressed size");
        }
        if ((flags & 0x80) != 0)
            declaredUncompressed = ReadVarint(_buffer, ref position, headerSize - 4, blockOffset, "the block uncompressed size");

        int deltaDistance = 0;
        uint dictionary = 0;
        for (int filter = 0; filter < filterCount; filter++) {
            long filterOffset = blockOffset + position;
            long id = ReadVarint(_buffer, ref position, headerSize - 4, blockOffset, "a filter ID");
            long propertiesSize = ReadVarint(_buffer, ref position, headerSize - 4, blockOffset, "the filter properties size");
            if (propertiesSize > headerSize - 4 - position) throw Corrupt(filterOffset, "filter properties inside the block header");
            bool last = filter == filterCount - 1;
            if (id == FilterLzma2) {
                if (!last) throw Corrupt(filterOffset, "LZMA2 only as the last filter");
                if (propertiesSize != 1 || _buffer[position] > 40) throw Corrupt(filterOffset, "one LZMA2 dictionary property byte of at most 40");
                dictionary = LzmaFolderDecoder.Lzma2DictionarySize(_buffer[position]);
            } else if (id == FilterDelta) {
                if (last) throw Corrupt(filterOffset, "an LZMA2 filter last in the chain");
                if (propertiesSize != 1 || deltaDistance != 0) throw Corrupt(filterOffset, "one delta filter with one property byte");
                deltaDistance = _buffer[position] + 1;
            } else if (id is >= 0x04 and <= 0x0B) {
                throw new NotSupportedException($"xz branch/call/jump filter 0x{id:x2} is unsupported.");
            } else {
                throw new NotSupportedException($"xz filter 0x{id:x} is unsupported.");
            }
            position += (int)propertiesSize;
        }
        for (; position < headerSize - 4; position++)
            if (_buffer[position] != 0) throw Corrupt(blockOffset + position, "null block header padding");
        if (dictionary > _settings.MaximumWindow)
            throw ArchiveLimits.Exceeded(nameof(ArchiveReadOptions.MaxWindowBytes), _settings.MaximumWindow);
        if (declaredUncompressed > _settings.MaximumOutput - _total)
            throw ArchiveLimits.Exceeded(_settings.OutputLimitName, _settings.MaximumOutput);

        long dataOffset = Offset;
        var sink = new XzBlockSink(_destination, check, deltaDistance, _settings, _total, declaredUncompressed, () => Corrupt(dataOffset,
            $"no more than the declared {declaredUncompressed} uncompressed bytes in the block"));
        var decoder = LzmaFolderDecoder.CreateStreaming(dictionary, sink);
        try {
            decoder.DecodeLzma2Streaming(_source, _settings.Budget, _settings.CancellationToken);
        } catch (InvalidDataException error) {
            throw new ArchiveCorruptionException(FormatName, dataOffset, "valid LZMA2 data in the block: " + error.Message, error);
        }
        long compressed = Offset - dataOffset;
        long uncompressed = sink.Written;
        if (declaredCompressed >= 0 && compressed != declaredCompressed)
            throw Corrupt(dataOffset, $"{declaredCompressed} bytes of compressed block data as declared, not {compressed}");
        if (declaredUncompressed >= 0 && uncompressed != declaredUncompressed)
            throw Corrupt(dataOffset, $"{declaredUncompressed} uncompressed bytes as declared, not {uncompressed}");

        long paddingOffset = Offset;
        int padding = (int)((4 - (headerSize + compressed) % 4) % 4);
        ReadExact(_buffer, padding, paddingOffset, "block padding");
        for (int index = 0; index < padding; index++)
            if (_buffer[index] != 0) throw Corrupt(paddingOffset + index, "null block padding");
        int checkSize = CheckSize(check);
        long checkOffset = Offset;
        ReadExact(_buffer, checkSize, checkOffset, "the block integrity check");
        if (!sink.CheckMatches(_buffer)) throw Corrupt(checkOffset, $"a {check} check that matches the block content");
        _total += uncompressed;
        return (headerSize + compressed + checkSize, uncompressed);
    }

    private long ReadIndex(long indexOffset, List<(long Unpadded, long Uncompressed)> records) {
        // The index indicator byte has already been read; its bytes are checked with the rest of the index.
        var index = new MemoryStream();
        index.WriteByte(0);
        long count = ReadStreamVarint(index, "the index record count");
        if (count != records.Count) throw Corrupt(indexOffset + 1, $"an index of {records.Count} records, not {count}");
        foreach ((long unpadded, long uncompressed) in records) {
            long recordOffset = Offset;
            long declaredUnpadded = ReadStreamVarint(index, "an index unpadded size");
            long declaredUncompressed = ReadStreamVarint(index, "an index uncompressed size");
            if (declaredUnpadded != unpadded || declaredUncompressed != uncompressed)
                throw Corrupt(recordOffset, $"an index record of {unpadded} unpadded and {uncompressed} uncompressed bytes");
        }
        while (index.Length % 4 != 0) {
            long paddingOffset = Offset;
            int value = ReadByte(paddingOffset, "index padding");
            if (value != 0) throw Corrupt(paddingOffset, "null index padding");
            index.WriteByte(0);
        }
        long crcOffset = Offset;
        ReadExact(_buffer, 4, crcOffset, "the index CRC32");
        byte[] bytes = index.ToArray();
        if (OfficePngCrc32.Compute(bytes, 0, bytes.Length) != ReadUInt32(_buffer, 0)) throw Corrupt(crcOffset, "an index CRC32 that matches");
        return bytes.Length + 4;
    }

    private long ReadStreamVarint(MemoryStream copy, string structure) {
        long start = Offset;
        long value = 0;
        for (int index = 0; index < 9; index++) {
            int octet = ReadByte(start, structure);
            copy.WriteByte((byte)octet);
            value |= (long)(octet & 0x7F) << (7 * index);
            if ((octet & 0x80) == 0) {
                if (octet == 0 && index > 0) throw Corrupt(start, "a minimally encoded " + structure);
                return value;
            }
        }
        throw Corrupt(start, structure + " of at most nine bytes");
    }

    private static long ReadVarint(byte[] data, ref int position, int end, long baseOffset, string structure) {
        int start = position;
        long value = 0;
        for (int index = 0; index < 9; index++) {
            if (position >= end) throw Corrupt(baseOffset + start, structure + " inside the block header");
            int octet = data[position++];
            value |= (long)(octet & 0x7F) << (7 * index);
            if ((octet & 0x80) == 0) {
                if (octet == 0 && index > 0) throw Corrupt(baseOffset + start, "a minimally encoded " + structure);
                return value;
            }
        }
        throw Corrupt(baseOffset + start, structure + " of at most nine bytes");
    }

    internal static int CheckSize(XzCheck check) => check switch {
        XzCheck.Crc32 => 4,
        XzCheck.Crc64 => 8,
        XzCheck.Sha256 => 32,
        _ => 0
    };

    private int ReadByte(long offset, string structure) {
        int value = _source.ReadByte();
        if (value < 0) throw Corrupt(offset, structure);
        return value;
    }

    private void ReadExact(byte[] buffer, int count, long offset, string structure) => ReadExact(buffer, 0, count, offset, structure);

    private void ReadExact(byte[] buffer, int start, int count, long offset, string structure) {
        int total = 0;
        while (total < count) {
            int read = _source.Read(buffer, start + total, count - total);
            if (read <= 0) throw Corrupt(offset, $"{count} bytes of {structure}, found {total}");
            total += read;
        }
    }

    private static uint ReadUInt32(byte[] data, int offset) =>
        (uint)(data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16 | data[offset + 3] << 24);
}

/// <summary>Writes one block's output through the optional delta filter and its integrity check.</summary>
internal sealed class XzBlockSink : LzmaOutputSink {
    private readonly Stream _destination;
    private readonly XzCheck _check;
    private readonly CompressedStreamSettings _settings;
    private readonly long _before;
    private readonly long _declared;
    private readonly Func<Exception> _overrun;
    private readonly byte[]? _delta;
    private readonly int _deltaDistance;
    private int _deltaPosition;
    private uint _crc32 = OfficePngCrc32.Begin();
    private ulong _crc64 = Crc64.Begin();
    private readonly IncrementalHash? _sha256;

    internal XzBlockSink(Stream destination, XzCheck check, int deltaDistance, CompressedStreamSettings settings, long before,
        long declared, Func<Exception> overrun) {
        _destination = destination;
        _check = check;
        _settings = settings;
        _before = before;
        _declared = declared;
        _overrun = overrun;
        _deltaDistance = deltaDistance;
        if (deltaDistance > 0) _delta = new byte[256];
        if (check == XzCheck.Sha256) _sha256 = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
    }

    internal override void Admit(long total) {
        if (_declared >= 0 && total > _declared) throw _overrun();
        if (total > _settings.MaximumOutput - _before) throw ArchiveLimits.Exceeded(_settings.OutputLimitName, _settings.MaximumOutput);
    }

    protected override void Consume(byte[] buffer, int offset, int count) {
        _settings.Budget.Charge(count, count);
        _settings.Select?.Invoke(count);
        if (_delta != null) {
            // out[i] = in[i] + out[i - distance], with the history starting as zeros. The decoder still needs its own
            // bytes as LZMA history, so the filtered bytes go to a separate buffer.
            var filtered = new byte[count];
            for (int index = 0; index < count; index++) {
                byte value = (byte)(buffer[offset + index] + _delta[(_deltaPosition - _deltaDistance) & 0xFF]);
                _delta[_deltaPosition & 0xFF] = value;
                _deltaPosition++;
                filtered[index] = value;
            }
            buffer = filtered;
            offset = 0;
        }
        switch (_check) {
            case XzCheck.Crc32:
                _crc32 = OfficePngCrc32.Append(_crc32, buffer, offset, count);
                break;
            case XzCheck.Crc64:
                _crc64 = Crc64.Append(_crc64, buffer, offset, count);
                break;
            case XzCheck.Sha256:
                _sha256!.AppendData(buffer, offset, count);
                break;
        }
        _destination.Write(buffer, offset, count);
    }

    internal bool CheckMatches(byte[] stored) {
        switch (_check) {
            case XzCheck.Crc32:
                uint crc32 = OfficePngCrc32.Complete(_crc32);
                return stored[0] == (byte)crc32 && stored[1] == (byte)(crc32 >> 8) && stored[2] == (byte)(crc32 >> 16) && stored[3] == (byte)(crc32 >> 24);
            case XzCheck.Crc64:
                ulong crc64 = Crc64.Complete(_crc64);
                for (int index = 0; index < 8; index++)
                    if (stored[index] != (byte)(crc64 >> (8 * index))) return false;
                return true;
            case XzCheck.Sha256:
                byte[] digest = _sha256!.GetHashAndReset();
                _sha256.Dispose();
                for (int index = 0; index < 32; index++)
                    if (stored[index] != digest[index]) return false;
                return true;
            default:
                return true;
        }
    }
}

internal readonly struct XzDecodeResult {
    internal XzDecodeResult(int streams, int blocks, XzCheck check, long length) {
        Streams = streams;
        Blocks = blocks;
        Check = check;
        Length = length;
    }

    internal int Streams { get; }
    internal int Blocks { get; }
    internal XzCheck Check { get; }
    internal long Length { get; }
}

/// <summary>Limits and accounting shared by the single-stream decoders.</summary>
internal sealed class CompressedStreamSettings {
    internal CompressedStreamSettings(long maximumOutput, string outputLimitName, long maximumWindow, ArchiveBlockBudget budget,
        Action<long>? select, CancellationToken cancellationToken) {
        MaximumOutput = maximumOutput;
        OutputLimitName = outputLimitName;
        MaximumWindow = maximumWindow;
        Budget = budget;
        Select = select;
        CancellationToken = cancellationToken;
    }

    internal long MaximumOutput { get; }
    internal string OutputLimitName { get; }
    internal long MaximumWindow { get; }
    internal ArchiveBlockBudget Budget { get; }
    internal Action<long>? Select { get; }
    internal CancellationToken CancellationToken { get; }
}
