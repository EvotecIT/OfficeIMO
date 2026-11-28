using OfficeIMO.Archives.Compression;

namespace OfficeIMO.Archives;

/// <summary>A single-stream compression wrapper that can hold a TAR archive or one file.</summary>
internal enum StreamCodec {
    None,
    GZip,
    Zstandard,
    Xz,
    Lzip
}

/// <summary>Detects, peeks into and decompresses the single-stream codecs that wrap TAR archives and single files.</summary>
internal static class StreamCodecs {
    /// <summary>Identifies a codec from its magic bytes; eight bytes are enough for every codec.</summary>
    internal static StreamCodec Detect(byte[] bytes, int length) {
        if (GZipContent.HasSignature(bytes, length)) return StreamCodec.GZip;
        if (length >= 4 && bytes[1] == 0xB5 && bytes[2] == 0x2F && bytes[3] == 0xFD && bytes[0] == 0x28) return StreamCodec.Zstandard;
        // A Zstandard skippable frame may come first: magic 0x184D2A50 to 0x184D2A5F, little-endian.
        if (length >= 4 && (bytes[0] & 0xF0) == 0x50 && bytes[1] == 0x2A && bytes[2] == 0x4D && bytes[3] == 0x18) return StreamCodec.Zstandard;
        if (length >= 6 && bytes[0] == 0xFD && bytes[1] == (byte)'7' && bytes[2] == (byte)'z' && bytes[3] == (byte)'X'
            && bytes[4] == (byte)'Z' && bytes[5] == 0) return StreamCodec.Xz;
        if (length >= 4 && bytes[0] == (byte)'L' && bytes[1] == (byte)'Z' && bytes[2] == (byte)'I' && bytes[3] == (byte)'P') return StreamCodec.Lzip;
        return StreamCodec.None;
    }

    /// <summary>The TAR-wrapped format for <paramref name="codec"/>.</summary>
    internal static ArchiveFormat TarFormat(StreamCodec codec) => codec switch {
        StreamCodec.GZip => ArchiveFormat.TarGZip,
        StreamCodec.Zstandard => ArchiveFormat.TarZstandard,
        StreamCodec.Xz => ArchiveFormat.TarXz,
        StreamCodec.Lzip => ArchiveFormat.TarLzip,
        _ => ArchiveFormat.Tar
    };

    /// <summary>The single-file format for <paramref name="codec"/>.</summary>
    internal static ArchiveFormat FileFormat(StreamCodec codec) => codec switch {
        StreamCodec.GZip => ArchiveFormat.GZip,
        StreamCodec.Zstandard => ArchiveFormat.Zstandard,
        StreamCodec.Xz => ArchiveFormat.Xz,
        StreamCodec.Lzip => ArchiveFormat.Lzip,
        _ => throw new ArgumentOutOfRangeException(nameof(codec))
    };

    /// <summary>The codec behind a TAR-wrapped or single-file format, or None.</summary>
    internal static StreamCodec Of(ArchiveFormat format) => format switch {
        ArchiveFormat.TarGZip or ArchiveFormat.GZip => StreamCodec.GZip,
        ArchiveFormat.TarZstandard or ArchiveFormat.Zstandard => StreamCodec.Zstandard,
        ArchiveFormat.TarXz or ArchiveFormat.Xz => StreamCodec.Xz,
        ArchiveFormat.TarLzip or ArchiveFormat.Lzip => StreamCodec.Lzip,
        _ => StreamCodec.None
    };

    /// <summary>File-name suffixes that a single compressed file conventionally carries.</summary>
    internal static string Suffix(StreamCodec codec) => codec switch {
        StreamCodec.GZip => ".gz",
        StreamCodec.Zstandard => ".zst",
        StreamCodec.Xz => ".xz",
        StreamCodec.Lzip => ".lz",
        _ => string.Empty
    };

    /// <summary>Decompresses everything from the current position of a seekable source. Output is bounded by
    /// <paramref name="maximumOutput"/>, reported under <paramref name="limitName"/>, and decode work is charged to the
    /// operation budget. With <paramref name="selected"/> the output is also charged as selected bytes.</summary>
    internal static void Decompress(StreamCodec codec, Stream source, Stream destination, long maximumOutput, ArchiveLimits limits,
        string limitName, bool selected, CancellationToken token) {
        Action<long>? select = selected ? bytes => limits.Operation.Select(bytes) : null;
        switch (codec) {
            case StreamCodec.GZip:
                GZipContent.Decompress(source, destination, maximumOutput, limits, limitName, token, selected);
                break;
            case StreamCodec.Zstandard:
                ZstandardDecoder.Decode(source, destination, new ZstandardDecoderSettings(maximumOutput, limitName, limits.Window,
                    new ArchiveBlockBudget(limits), select, null, token));
                break;
            case StreamCodec.Xz:
                XzDecoder.Decode(source, destination, Settings(maximumOutput, limitName, limits, select, token));
                break;
            case StreamCodec.Lzip:
                LzipDecoder.Decode(source, destination, Settings(maximumOutput, limitName, limits, select, token));
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(codec));
        }
    }

    /// <summary>Decompresses at most <paramref name="count"/> bytes for format detection, without charging the caller's
    /// budget. Damaged or truncated input returns what decoded before the problem.</summary>
    internal static int Peek(StreamCodec codec, Stream source, byte[] destination, int count, long maximumWindow, CancellationToken token) {
        if (codec == StreamCodec.GZip) return GZipContent.Peek(source, destination, count, token);
        var detached = new ArchiveLimits(new ArchiveExtractionOptions { MaxWindowBytes = maximumWindow });
        var sink = new PeekSink(destination, count);
        try {
            Decompress(codec, source, sink, long.MaxValue, detached, nameof(ArchiveExtractionOptions.MaxExpandedBytes), selected: false, token);
        } catch (PeekComplete) {
        } catch (Exception error) when (error is IOException or NotSupportedException or OfficeInspectionResourceLimitException) {
        }
        return sink.Filled;
    }

    private static CompressedStreamSettings Settings(long maximumOutput, string limitName, ArchiveLimits limits, Action<long>? select,
        CancellationToken token) =>
        new(maximumOutput, limitName, limits.Window, new ArchiveBlockBudget(limits), select, token);

    private sealed class PeekComplete : Exception {
    }

    /// <summary>Keeps the first bytes written and stops the decoder once it has them.</summary>
    private sealed class PeekSink : Stream {
        private readonly byte[] _destination;
        private readonly int _count;
        internal PeekSink(byte[] destination, int count) { _destination = destination; _count = count; }
        internal int Filled { get; private set; }
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => Filled;
        public override long Position { get => Filled; set => throw new NotSupportedException(); }
        public override void Write(byte[] buffer, int offset, int count) {
            int take = Math.Min(count, _count - Filled);
            Buffer.BlockCopy(buffer, offset, _destination, Filled, take);
            Filled += take;
            if (Filled == _count) throw new PeekComplete();
        }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }
}
