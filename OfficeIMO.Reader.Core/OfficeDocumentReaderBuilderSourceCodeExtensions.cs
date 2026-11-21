namespace OfficeIMO.Reader;

/// <summary>Adds dependency-free source-code ingestion to a Reader builder.</summary>
public static class OfficeDocumentReaderBuilderSourceCodeExtensions {
    /// <summary>Stable source-code handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.source-code";

    /// <summary>
    /// Adds bounded source-code ingestion for common programming, scripting and configuration file extensions.
    /// </summary>
    /// <remarks>
    /// Text is returned verbatim with normalized line breaks, Markdown wraps each chunk in a fenced block tagged with the
    /// language, and the language is reported as document metadata. Inputs that contain binary data are refused.
    /// Extensionless files named by convention, such as <c>Dockerfile</c>, <c>Makefile</c>, <c>Jenkinsfile</c> and
    /// <c>Gemfile</c>, are read as their language. Other extensionless scripts are recognised by their <c>#!</c> line
    /// when no other handler claims them.
    /// Extensions owned by structured readers, such as JSON, YAML, TOML, XML, Markdown, CSV and HTML, are not claimed.
    /// </remarks>
    public static OfficeDocumentReaderBuilder AddSourceCodeHandler(this OfficeDocumentReaderBuilder builder, bool replaceExisting = false) {
        ArgumentNullException.ThrowIfNull(builder);
        IReadOnlyList<string> extensions = ReaderSourceCodeLanguages.Extensions;
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "Source Code Reader",
            Description = "Dependency-free source-code ingestion with fenced Markdown and language metadata.",
            Kind = ReaderInputKind.Text,
            UseDetectedKindFallback = false,
            Extensions = extensions,
            FormatQualifications = extensions.Select(static extension => new ReaderFormatQualification(extension,
                "SourceCode" + extension, ReaderFormatSupport.ReadConvert, "BOM-aware text decoding with binary refusal",
                new[] { "Verbatim source text with normalized line breaks", "Fenced Markdown tagged with the language", "Language and line-count metadata" },
                new[] { "No syntax parsing, symbols or comments model", "Inputs containing NUL bytes outside UTF-16 or UTF-32 are refused" },
                new[] { "OfficeIMO.Reader.Tests/Reader.SourceCode.cs" })).ToArray(),
            SupportsTextAndMetadata = true,
            ReadPath = static (path, options, token) => SourceCodeReaderAdapter.ReadDocument(path, options, token).Chunks,
            ReadStream = static (stream, sourceName, options, token) => SourceCodeReaderAdapter.ReadDocument(stream, sourceName, options, token).Chunks,
            ReadDocumentPath = static (path, options, token) => SourceCodeReaderAdapter.ReadDocument(path, options, token),
            ReadDocumentStream = static (stream, sourceName, options, token) => SourceCodeReaderAdapter.ReadDocument(stream, sourceName, options, token),
            // Content detection can mistake source for Markdown or XML; a text payload keeps its source extension.
            ExtensionValidationProbeStream = static (stream, _, options, token) => SourceCodeReaderAdapter.Probe(stream, options, token),
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true
        }, replaceExisting);
    }
}
