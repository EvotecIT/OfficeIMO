using OfficeIMO.Core.Internal;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher;

/// <summary>A managed editable publication, loaded from Publisher or created in memory, with bounded native source saving and SVG export.</summary>
public sealed partial class PublisherDocument : IDisposable {
    private PublisherModelState? _model;
    internal PublisherDocument(IReadOnlyList<PublisherPage> pages, IReadOnlyList<PublisherPage> masters,
        IReadOnlyList<PublisherTextStory> stories, IReadOnlyList<PublisherImage> images, PublisherReadReport report, bool hasNativeSource = false) {
        Pages = Array.AsReadOnly(pages.ToArray()); MasterPages = Array.AsReadOnly(masters.ToArray());
        TextStories = Array.AsReadOnly(stories.ToArray()); Images = Array.AsReadOnly(images.ToArray()); ReadReport = report; SourceReadReport = report;
        HasNativeSource = hasNativeSource;
    }
    /// <summary>Whether this document was loaded from native publication bytes. In-memory edits and added pages do not change this provenance.</summary>
    public bool HasNativeSource { get; }
    /// <summary>Absolute path of the loaded file or most recent successful native path save. Stream operations do not associate a path.</summary>
    public string? FilePath { get; private set; }
    /// <summary>Document pages in current publication order. Utility pages and master definitions are excluded.</summary>
    public IReadOnlyList<PublisherPage> Pages { get; private set; }
    /// <summary>Current master definitions, also projected onto their referring document pages.</summary>
    public IReadOnlyList<PublisherPage> MasterPages { get; private set; }
    /// <summary>Complete current text stories, including authored stories and content whose placement could not be reconstructed.</summary>
    public IReadOnlyList<PublisherTextStory> TextStories { get; private set; }
    /// <summary>Recovered embedded image payloads.</summary>
    public IReadOnlyList<PublisherImage> Images { get; private set; }
    /// <summary>Source decoding and current page reconstruction report. Successful edits replace this snapshot with fresh projection evidence.</summary>
    public PublisherReadReport ReadReport { get; private set; }
    /// <summary>Immutable native decoding evidence captured before the initial page projection, or zero source evidence for source-free creation. Edits do not replace this snapshot.</summary>
    public PublisherReadReport SourceReadReport { get; private set; }
    /// <summary>Loads a publication from a file without retaining file handles.</summary>
    public static PublisherDocument Load(string path, PublisherLoadOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(path);
        cancellationToken.ThrowIfCancellationRequested();
        string fullPath = Path.GetFullPath(path);
        using var stream = File.OpenRead(fullPath);
        PublisherDocument document = Load(stream, options, cancellationToken);
        document.FilePath = fullPath;
        return document;
    }
    /// <summary>Loads bounded native publication bytes. Unsupported generations and malformed records are rejected.</summary>
    public static PublisherDocument Load(byte[] bytes, PublisherLoadOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(bytes);
        PublisherLoadOptions operation = (options ?? new PublisherLoadOptions()).Clone();
        cancellationToken.ThrowIfCancellationRequested();
        if (bytes.Length > operation.Limits.MaxInputBytes) throw new InvalidDataException("Publisher input byte limit exceeded.");
        var compoundOptions = new OfficeCompoundReadOptions(maxDirectoryEntries: operation.Limits.MaxRecords,
            maxStreamCount: operation.Limits.MaxCompoundStreams, maxStreamBytes: operation.Limits.MaxInputBytes,
            maxTotalStreamBytes: operation.Limits.MaxInputBytes);
        if (!OfficeCompoundFileReader.TryRead(bytes, compoundOptions, cancellationToken, out OfficeCompoundFile? compound, out string? error))
            throw new InvalidDataException(error ?? "Invalid Publisher compound document.");
        return PublisherNativeReader.Read(compound!, operation, cancellationToken, bytes);
    }
    /// <summary>Reads from the current position to EOF, leaving the caller's stream open. Non-seekable streams are supported.</summary>
    public static PublisherDocument Load(Stream stream, PublisherLoadOptions? options = null, CancellationToken cancellationToken = default) {
        PublisherLoadOptions operation = (options ?? new PublisherLoadOptions()).Clone();
        return Load(OfficeLegacyImportBuffer.ReadAll(stream, operation.Limits, cancellationToken), operation, cancellationToken);
    }

    /// <summary>Asynchronously reads bounded publication bytes from a file, then parses and projects them with cancellation. No file handle is retained.</summary>
    public static async Task<PublisherDocument> LoadAsync(string path, PublisherLoadOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(path);
        cancellationToken.ThrowIfCancellationRequested();
        string fullPath = Path.GetFullPath(path);
        using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, FileOptions.Asynchronous);
        PublisherDocument document = await LoadAsync(stream, options, cancellationToken).ConfigureAwait(false);
        document.FilePath = fullPath;
        return document;
    }

    /// <summary>Asynchronously reads bounded native bytes from the current position to EOF, then parses and projects them. Caller-owned and non-seekable streams remain open.</summary>
    public static async Task<PublisherDocument> LoadAsync(Stream stream, PublisherLoadOptions? options = null, CancellationToken cancellationToken = default) {
        PublisherLoadOptions operation = (options ?? new PublisherLoadOptions()).Clone();
        byte[] bytes;
        try {
            bytes = await OfficeStreamReader.ReadRemainingBytesAsync(stream, cancellationToken, operation.Limits.MaxInputBytes).ConfigureAwait(false);
        } catch (InvalidDataException exception) when (OfficeStreamReader.IsSizeLimitException(exception)) {
            throw new InvalidDataException($"Legacy input exceeds the configured {operation.Limits.MaxInputBytes} byte limit.");
        }
        return Load(bytes, operation, cancellationToken);
    }

    /// <summary>Completes the document lifetime without changing detached snapshots or disposing caller-owned codecs and font providers. No owned handles are retained.</summary>
    public void Dispose() { }
}
