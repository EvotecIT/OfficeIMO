using OfficeIMO.Reader.Html;

namespace OfficeIMO.Reader.Email;

/// <summary>Adds MIME HTML archive support to an OfficeIMO Reader builder.</summary>
public static class OfficeDocumentReaderBuilderMhtmlExtensions {
    /// <summary>Stable MHTML handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.mhtml";
    /// <summary>Default maximum MHTML input size accepted by the aggregate reader.</summary>
    public const long DefaultMaxInputBytes = 64L * 1024L * 1024L;

    /// <summary>Adds MHT and MHTML ingestion backed by OfficeIMO.Mhtml and the HTML projection.</summary>
    public static OfficeDocumentReaderBuilder AddMhtmlHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderHtmlOptions? htmlOptions = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderHtmlOptions? registeredOptions = ReaderHtmlOptionsCloner.CloneNullable(htmlOptions);
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "MHTML Reader",
            Description = "MIME HTML archive projection using OfficeIMO.Mhtml and OfficeIMO.Reader.Html.",
            Kind = ReaderInputKind.Html,
            UseDetectedKindFallback = false,
            Extensions = new[] { ".mht", ".mhtml" },
            DefaultMaxInputBytes = DefaultMaxInputBytes,
            ReadPath = (path, readerOptions, token) => MhtmlReaderAdapter.Read(path, readerOptions,
                ReaderHtmlOptionsCloner.CloneNullable(registeredOptions), token),
            ReadStream = (stream, sourceName, readerOptions, token) => MhtmlReaderAdapter.Read(stream, sourceName,
                readerOptions, ReaderHtmlOptionsCloner.CloneNullable(registeredOptions), token),
            ReadDocumentPath = (path, readerOptions, token) => MhtmlReaderAdapter.ReadDocument(path, readerOptions,
                ReaderHtmlOptionsCloner.CloneNullable(registeredOptions), token),
            ReadDocumentStream = (stream, sourceName, readerOptions, token) => MhtmlReaderAdapter.ReadDocument(stream,
                sourceName, readerOptions, ReaderHtmlOptionsCloner.CloneNullable(registeredOptions), token),
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true
        }, replaceExisting);
    }
}
