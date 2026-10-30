using OfficeIMO.Core.Internal;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher;

/// <summary>A managed recovery of a Microsoft Publisher publication. Native .pub writing is not provided.</summary>
public sealed partial class PublisherDocument {
    private PublisherModelState? _model;
    internal PublisherDocument(IReadOnlyList<PublisherPage> pages, IReadOnlyList<PublisherPage> masters,
        IReadOnlyList<PublisherTextStory> stories, IReadOnlyList<PublisherImage> images, PublisherReadReport report) {
        Pages = Array.AsReadOnly(pages.ToArray()); MasterPages = Array.AsReadOnly(masters.ToArray());
        TextStories = Array.AsReadOnly(stories.ToArray()); Images = Array.AsReadOnly(images.ToArray()); ReadReport = report; SourceReadReport = report;
    }
    /// <summary>Document pages in native publication order. Utility pages and master definitions are excluded.</summary>
    public IReadOnlyList<PublisherPage> Pages { get; private set; }
    /// <summary>Native master definitions, also projected onto their referring document pages.</summary>
    public IReadOnlyList<PublisherPage> MasterPages { get; private set; }
    /// <summary>Complete recovered text stories, including content whose placement could not be reconstructed.</summary>
    public IReadOnlyList<PublisherTextStory> TextStories { get; private set; }
    /// <summary>Recovered embedded image payloads.</summary>
    public IReadOnlyList<PublisherImage> Images { get; private set; }
    /// <summary>Source decoding and current page reconstruction report. Successful edits replace this snapshot with fresh projection evidence.</summary>
    public PublisherReadReport ReadReport { get; private set; }
    /// <summary>Immutable native decoding evidence captured before the initial page projection. Edits do not replace this source evidence.</summary>
    public PublisherReadReport SourceReadReport { get; private set; }
    /// <summary>Loads a publication from a file without retaining file handles.</summary>
    public static PublisherDocument Load(string path, PublisherReadOptions? options = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        cancellationToken.ThrowIfCancellationRequested();
        using var stream = File.OpenRead(path);
        return Load(stream, options, cancellationToken);
    }
    /// <summary>Loads bounded native publication bytes. Unsupported generations and malformed records are rejected.</summary>
    public static PublisherDocument Load(byte[] bytes, PublisherReadOptions? options = null, CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        PublisherReadOptions operation = (options ?? new PublisherReadOptions()).Clone();
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
    public static PublisherDocument Load(Stream stream, PublisherReadOptions? options = null, CancellationToken cancellationToken = default) {
        PublisherReadOptions operation = (options ?? new PublisherReadOptions()).Clone();
        return Load(OfficeLegacyImportBuffer.ReadAll(stream, operation.Limits, cancellationToken), operation, cancellationToken);
    }
}
