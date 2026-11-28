using OfficeIMO.Archives.Compression;

namespace OfficeIMO.Archives;

/// <summary>Outcome of decompressing a Zstandard stream.</summary>
public sealed class ZstandardReadResult {
    internal ZstandardReadResult(long frameCount, long skippableFrameCount, long length) {
        FrameCount = frameCount;
        SkippableFrameCount = skippableFrameCount;
        Length = length;
    }

    /// <summary>Zstandard frames decoded. Concatenated frames form one decompressed stream.</summary>
    public long FrameCount { get; }

    /// <summary>Skippable frames passed over without interpretation.</summary>
    public long SkippableFrameCount { get; }

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
    /// number are used as raw content, which must be at least 8 bytes.</summary>
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
    /// charged to the operation budget as selected bytes, and decode work by
    /// <see cref="ArchiveReadOptions.MaxDecompressionWorkBytes"/>. Declared windows are bounded by
    /// <see cref="ArchiveReadOptions.MaxWindowBytes"/>, except in single-segment frames, whose window is their declared
    /// content size and which the output limit bounds instead. The source stays open. Seekable input is bounded to the bytes
    /// present when decoding starts, and on success the source is left at the end of them. Nonseekable input is snapshotted
    /// through its end first, within the input limit. Output is written block by block, so a failure can leave a
    /// partial result in the destination. <paramref name="destination"/> must be a different stream from
    /// <paramref name="source"/> and must not write over the data being read.</remarks>
    /// <exception cref="ArchiveCorruptionException">The input is not Zstandard, is truncated, or fails a checksum.</exception>
    /// <exception cref="OfficeInspectionResourceLimitException">A configured limit is exceeded.</exception>
    public static ZstandardReadResult Decompress(Stream source, Stream destination, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) =>
        Decode(source, destination, null, options, cancellationToken);

    /// <summary>Decompresses frames that were compressed against <paramref name="dictionary"/>. A frame that names a
    /// different dictionary ID is rejected, and so is a frame that names any dictionary when
    /// <paramref name="dictionary"/> is raw content, which has no ID.</summary>
    /// <inheritdoc cref="Decompress(Stream, Stream, ArchiveReadOptions?, CancellationToken)"/>
    public static ZstandardReadResult DecompressWithDictionary(Stream source, Stream destination, ZstandardDictionary dictionary,
        ArchiveReadOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(dictionary);
        return Decode(source, destination, dictionary, options, cancellationToken);
    }

    private static ZstandardReadResult Decode(Stream source, Stream destination, ZstandardDictionary? dictionary,
        ArchiveReadOptions? options, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(destination);
        if (ReferenceEquals(source, destination)) throw new ArgumentException("The source and destination must be different streams.", nameof(destination));
        ArchiveReadOptions effective = options ?? new ArchiveReadOptions();
        if (effective.MaxWindowBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options), "MaxWindowBytes must be positive.");
        var limits = effective.Limits();
        using Stream? snapshot = source.CanSeek ? null : ArchiveReader.Snapshot(source, limits, cancellationToken);
        // Seekable input is read only within the range admitted here, so bytes appended later are never decoded.
        long start = source.CanSeek ? Math.Min(source.Position, source.Length) : 0;
        using Stream? admitted = snapshot == null ? Admit(source, start, limits) : null;
        Stream input = snapshot ?? admitted!;
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
        // The decoder skips over skippable frames by seeking the view, so leave the caller after everything it consumed.
        if (admitted != null) source.Position = start + admitted.Length;
        return new ZstandardReadResult(result.Frames, result.SkippableFrames, result.Length);
    }

    private static Stream Admit(Stream source, long start, ArchiveLimits limits) {
        // A position past the end, which MemoryStream allows, is empty input like a position at the end.
        long length = source.Length - start;
        ArchiveLimits.Check(length, limits.Input, nameof(ArchiveReadOptions.MaxInputBytes));
        limits.Operation.AdmitInput(length);
        return new OfficeIMO.Core.Internal.SeekableReadWindowStream(source, start, length);
    }
}
