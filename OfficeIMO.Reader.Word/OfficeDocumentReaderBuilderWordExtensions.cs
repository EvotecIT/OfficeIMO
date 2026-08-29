namespace OfficeIMO.Reader.Word;

/// <summary>Adds Word support to a modular Reader builder.</summary>
public static class OfficeDocumentReaderBuilderWordExtensions {
    /// <summary>Stable Word handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.word";
    /// <summary>Stable legacy-word handler identifier.</summary>
    public const string LegacyHandlerId = "officeimo.reader.word.legacy";

    /// <summary>Adds every Word format classified by <see cref="global::OfficeIMO.Word.WordFormatCatalog"/>.</summary>
    public static OfficeDocumentReaderBuilder AddWordHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderWordOptions? options = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderWordOptions configured = WordReaderAdapter.Clone(options);
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "Word Reader",
            Description = "OfficeIMO.Word Markdown and structured document projection.",
            Kind = ReaderInputKind.Word,
            Extensions = global::OfficeIMO.Word.WordFormatCatalog.All.Select(format => format.Extension).ToArray(),
            ReadDocumentPath = (path, readerOptions, token) => WordReaderAdapter.ReadDocument(path, readerOptions, configured, token),
            ReadDocumentStream = (stream, sourceName, readerOptions, token) => WordReaderAdapter.ReadDocument(stream, sourceName, readerOptions, configured, token),
            ProbeStream = (stream, sourceName, readerOptions, token) => WordReaderAdapter.ProbeEncryptedOpenXml(stream, readerOptions, token),
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true
        }, replaceExisting);
    }

    /// <summary>Adds safe read-only handlers for selected legacy word-processing families.</summary>
    public static OfficeDocumentReaderBuilder AddLegacyWordHandler(
        this OfficeDocumentReaderBuilder builder,
        global::OfficeIMO.Word.Legacy.LegacyWordImportOptions? importOptions = null,
        ReaderWordOptions? options = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderWordOptions configured = WordReaderAdapter.Clone(options);
        global::OfficeIMO.Word.Legacy.LegacyWordImportOptions? configuredImport = LegacyWordReaderAdapter.Clone(importOptions);
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = LegacyHandlerId,
            DisplayName = "Legacy Word Reader",
            Description = "Bounded WordPerfect, WordStar, Ami Pro, Word Pro, Works/Write, and Word for DOS import.",
            Kind = ReaderInputKind.Word,
            UseDetectedKindFallback = false,
            Extensions = new[] { ".wp", ".wp5", ".wp6", ".wpd", ".ws", ".ws3", ".ws4", ".ws5", ".ws6", ".ws7", ".sam", ".lwp", ".wps", ".wri" },
            ReadDocumentPath = (path, readerOptions, token) => LegacyWordReaderAdapter.ReadDocument(path, readerOptions, configured, configuredImport, token),
            ReadDocumentStream = (stream, sourceName, readerOptions, token) => LegacyWordReaderAdapter.ReadDocument(stream, sourceName, readerOptions, configured, configuredImport, token),
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true,
            DefaultMaxInputBytes = 64L * 1024L * 1024L
        }, replaceExisting);
    }
}
