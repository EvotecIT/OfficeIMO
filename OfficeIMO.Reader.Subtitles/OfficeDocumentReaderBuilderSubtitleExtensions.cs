namespace OfficeIMO.Reader.Subtitles;

/// <summary>Adds SRT and WebVTT support to <see cref="OfficeDocumentReaderBuilder"/>.</summary>
public static class OfficeDocumentReaderBuilderSubtitleExtensions {
    /// <summary>Stable handler identifier for subtitle adapter registration.</summary>
    public const string HandlerId = "officeimo.reader.subtitles";

    /// <summary>Default bounded subtitle size used when <see cref="ReaderOptions.MaxInputBytes"/> is not set.</summary>
    public const long DefaultMaxInputBytes = 32L * 1024L * 1024L;

    /// <summary>Adds bounded `.srt` and `.vtt` ingestion to an isolated reader builder.</summary>
    public static OfficeDocumentReaderBuilder AddSubtitleHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderSubtitleOptions? subtitleOptions = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderSubtitleOptions registered = (subtitleOptions ?? new ReaderSubtitleOptions()).CloneValidated();
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "Subtitle Reader Adapter",
            Description = "Bounded SRT and WebVTT transcript and cue-timing projection.",
            Kind = ReaderInputKind.Text,
            UseDetectedKindFallback = false,
            Extensions = new[] { ".srt", ".vtt" },
            DefaultMaxInputBytes = DefaultMaxInputBytes,
            ReadPath = (path, options, cancellationToken) => SubtitleReaderAdapter
                .ReadDocument(path, options, registered, cancellationToken).Chunks,
            ReadStream = (stream, sourceName, options, cancellationToken) => SubtitleReaderAdapter
                .ReadDocument(stream, sourceName, options, registered, cancellationToken).Chunks,
            ReadDocumentPath = (path, options, cancellationToken) => SubtitleReaderAdapter
                .ReadDocument(path, options, registered, cancellationToken),
            ReadDocumentStream = (stream, sourceName, options, cancellationToken) => SubtitleReaderAdapter
                .ReadDocument(stream, sourceName, options, registered, cancellationToken)
        }, replaceExisting);
    }
}
