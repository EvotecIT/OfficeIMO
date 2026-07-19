using OfficeIMO.PowerPoint.LegacyPpt;

namespace OfficeIMO.Reader.PowerPoint;

/// <summary>Adds PowerPoint support to a modular Reader builder.</summary>
public static class OfficeDocumentReaderBuilderPowerPointExtensions {
    /// <summary>Stable PowerPoint handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.powerpoint";

    /// <summary>Stable legacy binary PowerPoint handler identifier.</summary>
    public const string BinaryHandlerId = "officeimo.reader.powerpoint.binary";

    /// <summary>Adds PPTX, PPTM, and legacy PPT, POT, and PPS ingestion.</summary>
    public static OfficeDocumentReaderBuilder AddPowerPointHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderPowerPointOptions? options = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderPowerPointOptions configured = PowerPointReaderAdapter.Clone(options);
        builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "PowerPoint Reader",
            Description = "OfficeIMO.PowerPoint slide, table, and speaker-note projection.",
            Kind = ReaderInputKind.PowerPoint,
            Extensions = new[] { ".pptx", ".pptm" },
            ReadDocumentPath = (path, readerOptions, token) => PowerPointReaderAdapter.ReadDocument(path, readerOptions, configured, token),
            ReadDocumentStream = (stream, sourceName, readerOptions, token) => PowerPointReaderAdapter.ReadDocument(stream, sourceName, readerOptions, configured, token),
            ProbeStream = (stream, sourceName, readerOptions, token) => PowerPointReaderAdapter.ProbeEncryptedOpenXml(stream, readerOptions, token),
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true
        }, replaceExisting);
        builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = BinaryHandlerId,
            DisplayName = "Legacy PowerPoint Reader",
            Description = "Bounded OfficeIMO.PowerPoint projection for PPT, POT, and PPS compound files.",
            Kind = ReaderInputKind.PowerPoint,
            UseDetectedKindFallback = false,
            Extensions = new[] { ".ppt", ".pot", ".pps" },
            DefaultMaxInputBytes = LegacyPptImportOptions.DefaultMaxInputBytes,
            ReadDocumentPath = (path, readerOptions, token) => PowerPointReaderAdapter.ReadDocument(path, readerOptions, configured, token),
            ReadDocumentStream = (stream, sourceName, readerOptions, token) => PowerPointReaderAdapter.ReadDocument(stream, sourceName, readerOptions, configured, token),
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true
        }, replaceExisting);
        return builder;
    }
}
