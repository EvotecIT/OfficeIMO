using OfficeIMO.Archives.Compression;

namespace OfficeIMO.Archives;

/// <summary>Integrity checks an xz stream can declare.</summary>
public enum XzCheck {
    /// <summary>No integrity check.</summary>
    None = 0x00,
    /// <summary>CRC-32 of each block's uncompressed data.</summary>
    Crc32 = 0x01,
    /// <summary>CRC-64 (ECMA-182) of each block's uncompressed data.</summary>
    Crc64 = 0x04,
    /// <summary>SHA-256 of each block's uncompressed data.</summary>
    Sha256 = 0x0A
}

/// <summary>Outcome of decompressing an xz file.</summary>
public sealed class XzReadResult {
    internal XzReadResult(int streamCount, int blockCount, XzCheck check, long length) {
        StreamCount = streamCount;
        BlockCount = blockCount;
        Check = check;
        Length = length;
    }

    /// <summary>Concatenated xz streams decoded.</summary>
    public int StreamCount { get; }

    /// <summary>Blocks decoded across all streams.</summary>
    public int BlockCount { get; }

    /// <summary>Integrity check declared by the last stream.</summary>
    public XzCheck Check { get; }

    /// <summary>Decompressed bytes written to the destination.</summary>
    public long Length { get; }
}

/// <summary>Outcome of decompressing an lzip file.</summary>
public sealed class LzipReadResult {
    internal LzipReadResult(int memberCount, long length) {
        MemberCount = memberCount;
        Length = length;
    }

    /// <summary>Members decoded. Concatenated members form one decompressed stream.</summary>
    public int MemberCount { get; }

    /// <summary>Decompressed bytes written to the destination.</summary>
    public long Length { get; }
}

/// <summary>Decompresses xz files in managed code.</summary>
public static class XzArchive {
    /// <summary>Decompresses every stream from the current position of <paramref name="source"/> into
    /// <paramref name="destination"/>, verifying block checks (CRC-32, CRC-64 or SHA-256), declared sizes, the index and
    /// the stream footer. LZMA2 and the delta filter are supported.</summary>
    /// <remarks>Input is bounded by <see cref="ArchiveReadOptions.MaxInputBytes"/>; output by the smaller of
    /// <see cref="ArchiveReadOptions.MaxEntryBytes"/> and <see cref="ArchiveReadOptions.MaxExpandedBytes"/>, and it is
    /// charged to the operation budget as selected bytes. LZMA2 dictionaries are bounded by
    /// <see cref="ArchiveReadOptions.MaxWindowBytes"/>. Nonseekable input is snapshotted first. The source stays open.
    /// Output is written as it is decoded, so a failure can leave a partial result in the destination.</remarks>
    /// <exception cref="ArchiveCorruptionException">The input is not xz, is truncated, or fails a check.</exception>
    /// <exception cref="NotSupportedException">The input uses a branch/call/jump filter or an unsupported check type.</exception>
    /// <exception cref="OfficeInspectionResourceLimitException">A configured limit is exceeded.</exception>
    public static XzReadResult Decompress(Stream source, Stream destination, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        XzDecodeResult result = CompressedStreams.Run(source, destination, options, cancellationToken, XzDecoder.Decode);
        return new XzReadResult(result.Streams, result.Blocks, result.Check, result.Length);
    }
}

/// <summary>Decompresses lzip files in managed code.</summary>
public static class LzipArchive {
    /// <summary>Decompresses every member from the current position of <paramref name="source"/> into
    /// <paramref name="destination"/>, verifying each member's CRC-32, data size and member size. Data after the last
    /// member is rejected.</summary>
    /// <remarks>Limits apply as for <see cref="XzArchive.Decompress"/>; member dictionaries are bounded by
    /// <see cref="ArchiveReadOptions.MaxWindowBytes"/>.</remarks>
    /// <exception cref="ArchiveCorruptionException">The input is not lzip, is truncated, or fails a check.</exception>
    /// <exception cref="OfficeInspectionResourceLimitException">A configured limit is exceeded.</exception>
    public static LzipReadResult Decompress(Stream source, Stream destination, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        (int members, long length) = CompressedStreams.Run(source, destination, options, cancellationToken, LzipDecoder.Decode);
        return new LzipReadResult(members, length);
    }
}

internal static class CompressedStreams {
    internal static T Run<T>(Stream source, Stream destination, ArchiveReadOptions? options, CancellationToken cancellationToken,
        Func<Stream, Stream, CompressedStreamSettings, T> decode) {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(destination);
        ArchiveReadOptions effective = options ?? new ArchiveReadOptions();
        if (effective.MaxWindowBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options), "MaxWindowBytes must be positive.");
        var limits = effective.Limits();
        using Stream? snapshot = source.CanSeek ? null : ArchiveReader.Snapshot(source, limits, cancellationToken);
        Stream input = snapshot ?? source;
        if (snapshot == null) {
            ArchiveLimits.Check(input.Length - input.Position, limits.Input, nameof(ArchiveReadOptions.MaxInputBytes));
            limits.Operation.AdmitInput(input.Length - input.Position);
        }
        bool entryLimit = limits.Entry <= limits.Expanded;
        var settings = new CompressedStreamSettings(
            entryLimit ? limits.Entry : limits.Expanded,
            entryLimit ? nameof(ArchiveReadOptions.MaxEntryBytes) : nameof(ArchiveReadOptions.MaxExpandedBytes),
            effective.MaxWindowBytes,
            new ArchiveBlockBudget(limits),
            bytes => limits.Operation.Select(bytes),
            cancellationToken);
        return decode(input, destination, settings);
    }
}
