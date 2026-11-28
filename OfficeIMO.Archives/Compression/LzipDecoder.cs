using OfficeIMO.Archives.SevenZip;
using OfficeIMO.Drawing;

namespace OfficeIMO.Archives.Compression;

/// <summary>Decodes lzip files (lzip manual, "File format"): one or more members, each an LZMA stream with fixed
/// literal and position properties, an end marker and a trailer with CRC-32, data size and member size.</summary>
internal sealed class LzipDecoder {
    internal const string FormatName = "lzip";
    private const int HeaderSize = 6;
    private const int TrailerSize = 20;
    // lc = 3, lp = 0, pb = 2.
    private const int LzmaProperties = (2 * 5 + 0) * 9 + 3;

    private readonly Stream _source;
    private readonly Stream _destination;
    private readonly CompressedStreamSettings _settings;
    private readonly long _start;
    private readonly byte[] _buffer = new byte[TrailerSize];
    private long _total;

    private LzipDecoder(Stream source, Stream destination, CompressedStreamSettings settings) {
        _source = source;
        _destination = destination;
        _settings = settings;
        _start = source.Position;
    }

    internal static (int Members, long Length) Decode(Stream source, Stream destination, CompressedStreamSettings settings) =>
        new LzipDecoder(source, destination, settings).DecodeAll();

    private long Offset => _source.Position - _start;

    private static ArchiveCorruptionException Corrupt(long offset, string expected) => new(FormatName, offset, expected);

    private (int, long) DecodeAll() {
        int members = 0;
        if (_source.Position == _source.Length) throw Corrupt(0, "an lzip member header");
        while (_source.Position < _source.Length) {
            _settings.CancellationToken.ThrowIfCancellationRequested();
            DecodeMember(members == 0);
            members++;
        }
        return (members, _total);
    }

    private void DecodeMember(bool first) {
        long memberOffset = Offset;
        ReadExact(HeaderSize, memberOffset, "the 6-byte member header");
        if (_buffer[0] != 'L' || _buffer[1] != 'Z' || _buffer[2] != 'I' || _buffer[3] != 'P')
            throw Corrupt(memberOffset, first ? "the lzip magic bytes LZIP" : "another lzip member and no trailing data");
        if (_buffer[4] != 1) throw Corrupt(memberOffset + 4, "lzip version 1");
        int coded = _buffer[5];
        int log = coded & 0x1F;
        if (log < 12 || log > 29) throw Corrupt(memberOffset + 5, "a dictionary size between 4 KiB and 512 MiB");
        long dictionary = (1L << log) - ((1L << log) >> 4) * (coded >> 5);
        if (dictionary < 4096) throw Corrupt(memberOffset + 5, "a dictionary size between 4 KiB and 512 MiB");
        if (dictionary > _settings.MaximumWindow)
            throw ArchiveLimits.Exceeded(nameof(ArchiveReadOptions.MaxWindowBytes), _settings.MaximumWindow);

        long dataOffset = Offset;
        var sink = new LzipSink(_destination, _settings, _total);
        var decoder = LzmaFolderDecoder.CreateStreaming((uint)dictionary, sink);
        try {
            decoder.DecodeLzmaUntilEndMarker(_source, LzmaProperties, _settings.Budget, _settings.CancellationToken);
        } catch (InvalidDataException error) {
            throw new ArchiveCorruptionException(FormatName, dataOffset, "valid LZMA data ending with an end marker: " + error.Message, error);
        }

        long trailerOffset = Offset;
        ReadExact(TrailerSize, trailerOffset, "the 20-byte member trailer");
        uint crc = (uint)(_buffer[0] | _buffer[1] << 8 | _buffer[2] << 16 | _buffer[3] << 24);
        long dataSize = ReadInt64(4);
        long memberSize = ReadInt64(12);
        if (crc != sink.Crc) throw Corrupt(trailerOffset, "a CRC32 that matches the member data");
        if (dataSize != sink.Written) throw Corrupt(trailerOffset + 4, $"a data size of {sink.Written}");
        long actualMemberSize = Offset - memberOffset;
        if (memberSize != actualMemberSize) throw Corrupt(trailerOffset + 12, $"a member size of {actualMemberSize}");
        _total += sink.Written;
    }

    private long ReadInt64(int offset) {
        ulong value = 0;
        for (int index = 0; index < 8; index++) value |= (ulong)_buffer[offset + index] << (8 * index);
        return value > long.MaxValue ? -1 : (long)value;
    }

    private void ReadExact(int count, long offset, string structure) {
        int total = 0;
        while (total < count) {
            int read = _source.Read(_buffer, total, count - total);
            if (read <= 0) throw Corrupt(offset, $"{count} bytes of {structure}, found {total}");
            total += read;
        }
    }

    private sealed class LzipSink : LzmaOutputSink {
        private readonly Stream _destination;
        private readonly CompressedStreamSettings _settings;
        private readonly long _before;
        private uint _crc = OfficePngCrc32.Begin();

        internal LzipSink(Stream destination, CompressedStreamSettings settings, long before) {
            _destination = destination;
            _settings = settings;
            _before = before;
        }

        internal uint Crc => OfficePngCrc32.Complete(_crc);

        internal override void Admit(long total) {
            if (total > _settings.MaximumOutput - _before) throw ArchiveLimits.Exceeded(_settings.OutputLimitName, _settings.MaximumOutput);
        }

        protected override void Consume(byte[] buffer, int offset, int count) {
            _settings.Budget.Charge(count, count);
            _settings.Select?.Invoke(count);
            _crc = OfficePngCrc32.Append(_crc, buffer, offset, count);
            _destination.Write(buffer, offset, count);
        }
    }
}
