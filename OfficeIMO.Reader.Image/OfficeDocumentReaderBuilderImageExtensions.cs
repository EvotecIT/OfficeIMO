namespace OfficeIMO.Reader.Image;

/// <summary>Adds standalone image support to <see cref="OfficeDocumentReaderBuilder"/>.</summary>
public static class OfficeDocumentReaderBuilderImageExtensions {
    /// <summary>Stable handler identifier for image adapter registration.</summary>
    public const string HandlerId = "officeimo.reader.image";

    /// <summary>Default bounded image size used when <see cref="ReaderOptions.MaxInputBytes"/> is not set.</summary>
    public const long DefaultMaxInputBytes = 128L * 1024L * 1024L;

    /// <summary>Adds header-only image metadata, asset, and OCR-readiness projection.</summary>
    public static OfficeDocumentReaderBuilder AddImageHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderImageOptions? imageOptions = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderImageOptions registered = (imageOptions ?? new ReaderImageOptions()).CloneValidated();
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "Image Reader Adapter",
            Description = "Header-only image metadata, materializable asset, and optional OCR candidate projection.",
            Kind = ReaderInputKind.Unknown,
            UseDetectedKindFallback = false,
            Extensions = ImageReaderAdapter.Extensions,
            DefaultMaxInputBytes = DefaultMaxInputBytes,
            ReadPath = (path, options, cancellationToken) => ImageReaderAdapter
                .ReadDocument(path, options, registered, cancellationToken).Chunks,
            ReadStream = (stream, sourceName, options, cancellationToken) => ImageReaderAdapter
                .ReadDocument(stream, sourceName, options, registered, cancellationToken).Chunks,
            ReadDocumentPath = (path, options, cancellationToken) => ImageReaderAdapter
                .ReadDocument(path, options, registered, cancellationToken),
            ReadDocumentStream = (stream, sourceName, options, cancellationToken) => ImageReaderAdapter
                .ReadDocument(stream, sourceName, options, registered, cancellationToken)
        }, replaceExisting);
    }
}
