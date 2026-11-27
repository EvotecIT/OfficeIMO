using Microsoft.Extensions.DocumentExtraction;
using Microsoft.Extensions.Documents;

namespace OfficeIMO.Reader.DocumentExtraction;

/// <summary>
/// A <see cref="DocumentExtractionClient"/> that extracts documents in process with an <see cref="OfficeDocumentReader"/>.
/// </summary>
/// <remarks>
/// The client reads with whatever handlers the reader was built with. Add <c>OfficeIMO.Reader.All</c> and call
/// <c>AddAllOfficeIMOHandlers()</c> on the <see cref="OfficeDocumentReaderBuilder"/> to cover every OfficeIMO format,
/// or register only the handler packages you need. The declared media type selects the handler when it implies a file
/// extension; otherwise the identifier's extension and content detection decide.
/// </remarks>
public sealed class OfficeIMODocumentExtractionClient : DocumentExtractionClient {
    private readonly OfficeDocumentReader _reader;
    private readonly OfficeIMODocumentExtractionOptions _options;

    /// <summary>Creates a client over an already configured reader.</summary>
    /// <param name="reader">The reader whose handlers decide which formats can be extracted.</param>
    /// <param name="options">Optional client defaults. The values are copied when the client is created.</param>
    public OfficeIMODocumentExtractionClient(OfficeDocumentReader reader, OfficeIMODocumentExtractionOptions? options = null) {
        ArgumentNullException.ThrowIfNull(reader);
        _reader = reader;
        _options = (options ?? new OfficeIMODocumentExtractionOptions()).Clone();
    }

    /// <summary>Creates a client over a reader configured by <paramref name="configureReader"/>.</summary>
    /// <param name="configureReader">Registers reader handlers, for example <c>b =&gt; b.AddAllOfficeIMOHandlers()</c>.</param>
    /// <param name="options">Optional client defaults. The values are copied when the client is created.</param>
    public OfficeIMODocumentExtractionClient(Action<OfficeDocumentReaderBuilder> configureReader, OfficeIMODocumentExtractionOptions? options = null)
        : this(BuildReader(configureReader), options) {
    }

    /// <summary>The reader used for extraction.</summary>
    public OfficeDocumentReader Reader => _reader;

    /// <inheritdoc />
    /// <remarks>
    /// <see cref="DocumentExtractionOptions.ExtractorName"/> is not supported and must be null. The only supported
    /// additional property is <see cref="OfficeIMOExtractionPropertyNames.ReaderOptionsOverride"/>, whose value must be a
    /// <see cref="ReaderOptions"/> instance. The resulting document's <c>RawRepresentation</c> is the
    /// <see cref="OfficeDocumentReadResult"/>. An empty <paramref name="mediaType"/> leaves handler selection to the
    /// identifier's extension and content detection.
    /// </remarks>
    public override async Task<ExtractionDocument> ExtractAsync(
        Stream document,
        string identifier,
        string mediaType,
        DocumentExtractionOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(identifier);
        cancellationToken.ThrowIfCancellationRequested();

        // Snapshot the per-call settings before any asynchronous work, as the extraction contract requires.
        ReaderOptions? readerOptions = ResolveReaderOptions(options);
        OfficeIMODocumentExtractionOptions mapping = _options.Clone();
        string sourceName = OfficeIMOMediaTypes.ResolveSourceName(identifier, mediaType);

        OfficeDocumentReadResult result = await _reader
            .ReadDocumentAsync(document, sourceName, readerOptions, cancellationToken)
            .ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return ExtractionDocumentMapper.Map(result, identifier, mapping);
    }

    /// <inheritdoc />
    public override object? GetService(Type serviceType, object? serviceKey = null) {
        ArgumentNullException.ThrowIfNull(serviceType);
        if (serviceKey != null) return null;
        if (serviceType.IsInstanceOfType(this)) return this;
        if (serviceType.IsInstanceOfType(_reader)) return _reader;
        return null;
    }

    private ReaderOptions? ResolveReaderOptions(DocumentExtractionOptions? options) {
        if (options == null) return _options.ReaderOptions;
        if (options.ExtractorName != null) {
            throw new NotSupportedException(
                "The OfficeIMO extraction client has no named extractors. Select formats by registering reader handlers instead.");
        }

        ReaderOptions? readerOptions = _options.ReaderOptions;
        foreach (KeyValuePair<string, object?> property in options.AdditionalProperties) {
            if (!string.Equals(property.Key, OfficeIMOExtractionPropertyNames.ReaderOptionsOverride, StringComparison.Ordinal)) {
                throw new NotSupportedException("The OfficeIMO extraction client does not support the option '" + property.Key + "'.");
            }
            readerOptions = property.Value as ReaderOptions ?? throw new ArgumentException(
                "The '" + OfficeIMOExtractionPropertyNames.ReaderOptionsOverride + "' option must be a " + nameof(ReaderOptions) + " instance.",
                nameof(options));
        }
        return readerOptions;
    }

    private static OfficeDocumentReader BuildReader(Action<OfficeDocumentReaderBuilder> configureReader) {
        ArgumentNullException.ThrowIfNull(configureReader);
        var builder = new OfficeDocumentReaderBuilder();
        configureReader(builder);
        return builder.Build();
    }
}
