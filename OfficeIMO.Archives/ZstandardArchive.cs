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
    /// number are used as raw content, which must be at least 8 bytes. The content is held to
    /// <see cref="ArchiveReadOptions.MaxWindowBytes"/> of <paramref name="options"/> before it is copied.</summary>
    /// <exception cref="ArchiveCorruptionException">The dictionary header or tables are malformed.</exception>
    /// <exception cref="OfficeInspectionResourceLimitException">The content is larger than the window limit.</exception>
    public static ZstandardDictionary Load(byte[] dictionary, ArchiveReadOptions? options = null) {
        ArgumentNullException.ThrowIfNull(dictionary);
        return new ZstandardDictionary(ZstandardDictionaryData.Parse(dictionary, MaxContent(options)));
    }

    /// <summary>Uses <paramref name="content"/> as a raw-content dictionary, even when it starts with the dictionary magic
    /// number, as <c>zstd</c>'s raw dictionary mode does. Raw content must be at least 8 bytes and has dictionary ID zero.
    /// It is held to <see cref="ArchiveReadOptions.MaxWindowBytes"/> of <paramref name="options"/> before it is copied.</summary>
    /// <exception cref="ArchiveCorruptionException"><paramref name="content"/> is shorter than 8 bytes.</exception>
    /// <exception cref="OfficeInspectionResourceLimitException">The content is larger than the window limit.</exception>
    public static ZstandardDictionary LoadRaw(byte[] content, ArchiveReadOptions? options = null) {
        ArgumentNullException.ThrowIfNull(content);
        return new ZstandardDictionary(ZstandardDictionaryData.Raw(content, MaxContent(options)));
    }

    private static long MaxContent(ArchiveReadOptions? options) {
        long maximum = options?.MaxWindowBytes ?? ArchiveReadOptions.DefaultMaxWindowBytes;
        if (maximum <= 0) throw new ArgumentOutOfRangeException(nameof(options), "MaxWindowBytes must be positive.");
        return maximum;
    }
}

/// <summary>Decompresses Zstandard (RFC 8878) streams in managed code.</summary>
public static class ZstandardArchive {
    /// <summary>Decompresses every frame from the current position of <paramref name="source"/> into
    /// <paramref name="destination"/>, passing over skippable frames and verifying content sizes and checksums.</summary>
    /// <remarks>Input is bounded by <see cref="ArchiveReadOptions.MaxInputBytes"/>; output by the smaller of
    /// <see cref="ArchiveReadOptions.MaxEntryBytes"/> and <see cref="ArchiveReadOptions.MaxExpandedBytes"/>, and it is
    /// charged to the operation budget as selected bytes, and decode work by
    /// <see cref="ArchiveReadOptions.MaxDecompressionWorkBytes"/>. Windows, including the declared content
    /// size of a single-segment frame, and dictionary content are bounded by <see cref="ArchiveReadOptions.MaxWindowBytes"/>.
    /// The source stays open. Seekable input is bounded to the bytes
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
        if (!source.CanRead) throw new ArgumentException("The source must be readable.", nameof(source));
        if (!destination.CanWrite) throw new ArgumentException("The destination must be writable.", nameof(destination));
        if (ReferenceEquals(source, destination)) throw new ArgumentException("The source and destination must be different streams.", nameof(destination));
        ArchiveReadOptions effective = options ?? new ArchiveReadOptions();
        // Read once, so a callback that changes the options mid-decode cannot split validation from use.
        long maxWindow = effective.MaxWindowBytes;
        if (maxWindow <= 0) throw new ArgumentOutOfRangeException(nameof(options), "MaxWindowBytes must be positive.");
        // A call that is already cancelled admits nothing, so it leaves a shared operation budget untouched.
        cancellationToken.ThrowIfCancellationRequested();
        // Dictionary content becomes history in every frame, so it is admitted against the window limit up front.
        if (dictionary != null && dictionary.ContentLength > maxWindow)
            throw ArchiveLimits.Exceeded(nameof(ArchiveReadOptions.MaxWindowBytes), maxWindow);
        var limits = effective.Limits();
        using Stream? snapshot = source.CanSeek ? null : ArchiveReader.Snapshot(source, limits, cancellationToken);
        // Seekable input is read only within the range admitted here, so bytes appended later are never decoded.
        // The length is read once, so the start and the admitted range agree even if the source changes meanwhile.
        long end = source.CanSeek ? source.Length : 0;
        // A position past the end, which MemoryStream allows, is empty input like a position at the end.
        long start = source.CanSeek ? Math.Min(source.Position, end) : 0;
        using Stream? admitted = snapshot == null ? Admit(source, start, end - start, limits, cancellationToken) : null;
        Stream input = snapshot ?? admitted!;
        bool entryLimit = limits.Entry <= limits.Expanded;
        var settings = new ZstandardDecoderSettings(
            entryLimit ? limits.Entry : limits.Expanded,
            entryLimit ? nameof(ArchiveReadOptions.MaxEntryBytes) : nameof(ArchiveReadOptions.MaxExpandedBytes),
            maxWindow,
            new ArchiveBlockBudget(limits),
            bytes => limits.Operation.Select(bytes),
            dictionary?.Data,
            cancellationToken);
        ZstandardDecodeResult result = ZstandardDecoder.Decode(input, destination, settings);
        // The decoder skips over skippable frames by seeking the view, so leave the caller after everything it consumed.
        if (admitted != null) source.Position = start + admitted.Length;
        return new ZstandardReadResult(result.Frames, result.SkippableFrames, result.Length);
    }

    private static Stream Admit(Stream source, long start, long length, ArchiveLimits limits, CancellationToken cancellationToken) {
        // Measuring the source can request cancellation; report that rather than a limit the measured length would break.
        cancellationToken.ThrowIfCancellationRequested();
        ArchiveLimits.Check(length, limits.Input, nameof(ArchiveReadOptions.MaxInputBytes));
        Stream view;
        try {
            // Opening the view validates the range, so a source that shrank since it was measured is charged nothing.
            view = new OfficeIMO.Core.Internal.SeekableReadWindowStream(source, start, length, cancellationToken: cancellationToken);
        } catch (ArgumentOutOfRangeException) {
            // The source shrank between measuring it and opening the view.
            throw new ArchiveCorruptionException(ZstandardDecoder.FormatName, 0, $"the {length} bytes of input measured when decoding started");
        }
        try {
            // Measuring the source can take long enough for cancellation to arrive; nothing is charged after it does.
            cancellationToken.ThrowIfCancellationRequested();
            limits.Operation.AdmitInput(length);
            return view;
        } catch {
            view.Dispose();
            throw;
        }
    }
}
