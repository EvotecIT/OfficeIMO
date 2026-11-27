using OfficeIMO.Archives.Compression;

namespace OfficeIMO.Archives;

/// <summary>Outcome of decompressing a Zstandard stream.</summary>
public sealed class ZstandardReadResult {
    internal ZstandardReadResult(int frameCount, int skippableFrameCount, long length) {
        FrameCount = frameCount;
        SkippableFrameCount = skippableFrameCount;
        Length = length;
    }

    /// <summary>Zstandard frames decoded. Concatenated frames form one decompressed stream.</summary>
    public int FrameCount { get; }

    /// <summary>Skippable frames passed over without interpretation.</summary>
    public int SkippableFrameCount { get; }

    /// <summary>Decompressed bytes written to the destination.</summary>
    public long Length { get; }
}

/// <summary>A Zstandard dictionary for frames compressed against shared content (RFC 8878, section 5).</summary>
public sealed class ZstandardDictionary {
    private ZstandardDictionary(ZstandardDictionaryData data) => Data = data;

    internal ZstandardDictionaryData Data { get; }

    /// <summary>Dictionary ID from the dictionary header, or zero for a raw-content dictionary.</summary>
    public uint Id => Data.Id;

    /// <summary>True when the dictionary carries its own entropy tables and repeat offsets.</summary>
    public bool HasEntropyTables => Data.HasEntropyTables;

    /// <summary>Length of the dictionary content that frames may reference.</summary>
    public int ContentLength => Data.Content.Length;

    /// <summary>Parses a dictionary, such as one written by <c>zstd --train</c>. Bytes without the dictionary magic
    /// number are used as raw content.</summary>
    /// <exception cref="ArchiveCorruptionException">The dictionary header or tables are malformed.</exception>
    public static ZstandardDictionary Load(byte[] dictionary) {
        ArgumentNullException.ThrowIfNull(dictionary);
        return new ZstandardDictionary(ZstandardDictionaryData.Parse(dictionary));
    }
}

/// <summary>Decompresses Zstandard (RFC 8878) streams in managed code.</summary>
public static class ZstandardArchive {
    /// <summary>Decompresses every frame from the current position of <paramref name="source"/> into
    /// <paramref name="destination"/>, passing over skippable frames and verifying content sizes and checksums.</summary>
    /// <remarks>Input is bounded by <see cref="ArchiveReadOptions.MaxInputBytes"/>; output by the smaller of
    /// <see cref="ArchiveReadOptions.MaxEntryBytes"/> and <see cref="ArchiveReadOptions.MaxExpandedBytes"/>, and it is
    /// charged to the operation budget as selected bytes. Declared windows are bounded by
    /// <see cref="ArchiveReadOptions.MaxWindowBytes"/>. Nonseekable input is snapshotted first. The source stays open.
    /// Output is written block by block, so a failure can leave a partial result in the destination.</remarks>
    /// <exception cref="ArchiveCorruptionException">The input is not Zstandard, is truncated, or fails a checksum.</exception>
    /// <exception cref="OfficeInspectionResourceLimitException">A configured limit is exceeded.</exception>
    public static ZstandardReadResult Decompress(Stream source, Stream destination, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) =>
        Decompress(source, destination, null, options, cancellationToken);

    /// <summary>Decompresses frames that were compressed against <paramref name="dictionary"/>. A frame that names a
    /// different dictionary ID is rejected.</summary>
    /// <inheritdoc cref="Decompress(Stream, Stream, ArchiveReadOptions?, CancellationToken)"/>
    public static ZstandardReadResult Decompress(Stream source, Stream destination, ZstandardDictionary? dictionary,
        ArchiveReadOptions? options = null, CancellationToken cancellationToken = default) {
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
        var settings = new ZstandardDecoderSettings(
            entryLimit ? limits.Entry : limits.Expanded,
            entryLimit ? nameof(ArchiveReadOptions.MaxEntryBytes) : nameof(ArchiveReadOptions.MaxExpandedBytes),
            effective.MaxWindowBytes,
            new ArchiveBlockBudget(limits),
            bytes => limits.Operation.Select(bytes),
            dictionary?.Data,
            cancellationToken);
        ZstandardDecodeResult result = ZstandardDecoder.Decode(input, destination, settings);
        return new ZstandardReadResult(result.Frames, result.SkippableFrames, result.Length);
    }
}
