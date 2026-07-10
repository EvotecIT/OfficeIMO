namespace OfficeIMO.AsciiDoc;

/// <summary>Explicit preprocessing options for attributes, conditionals, and includes.</summary>
public sealed class AsciiDocProcessorOptions {
    /// <summary>Initial document attributes.</summary>
    public IReadOnlyDictionary<string, string>? Attributes { get; set; }

    /// <summary>Optional source identifier or primary file path.</summary>
    public string? SourceName { get; set; }

    /// <summary>Include resolver. Null disables all include reads.</summary>
    public IAsciiDocIncludeResolver? IncludeResolver { get; set; }

    /// <summary>Maximum nested includes.</summary>
    public int MaximumIncludeDepth { get; set; } = 16;

    /// <summary>Maximum resolved include files.</summary>
    public int MaximumIncludeCount { get; set; } = 256;

    /// <summary>Maximum total characters returned by include resolvers.</summary>
    public int MaximumIncludedCharacters { get; set; } = 16 * 1024 * 1024;

    /// <summary>Maximum characters in the processed source.</summary>
    public int MaximumOutputLength { get; set; } = 64 * 1024 * 1024;

    /// <summary>Attribute expansion behavior.</summary>
    public AsciiDocUndefinedAttributeBehavior UndefinedAttributeBehavior { get; set; } = AsciiDocUndefinedAttributeBehavior.Preserve;

    /// <summary>Explicit custom processors. Null means no extensions.</summary>
    public AsciiDocExtensionRegistry? Extensions { get; set; }

    /// <summary>Maximum custom directive invocations.</summary>
    public int MaximumExtensionInvocations { get; set; } = 256;
}

/// <summary>Diagnostic from explicit preprocessing.</summary>
public sealed class AsciiDocProcessingDiagnostic {
    internal AsciiDocProcessingDiagnostic(
        string code,
        AsciiDocDiagnosticSeverity severity,
        string message,
        string? sourceName,
        int line) {
        Code = code;
        Severity = severity;
        Message = message;
        SourceName = sourceName;
        Line = line;
    }

    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Severity.</summary>
    public AsciiDocDiagnosticSeverity Severity { get; }

    /// <summary>Message.</summary>
    public string Message { get; }

    /// <summary>Source identifier, when known.</summary>
    public string? SourceName { get; }

    /// <summary>One-based source line.</summary>
    public int Line { get; }
}

/// <summary>Original and explicitly processed AsciiDoc views.</summary>
public sealed class AsciiDocProcessingResult {
    internal AsciiDocProcessingResult(
        AsciiDocDocument sourceDocument,
        AsciiDocDocument document,
        string processedSource,
        AsciiDocDocumentAttributes attributes,
        IReadOnlyList<AsciiDocProcessingDiagnostic> diagnostics) {
        SourceDocument = sourceDocument;
        Document = document;
        ProcessedSource = processedSource;
        Attributes = attributes;
        Diagnostics = diagnostics;
    }

    /// <summary>Lossless unprocessed source document.</summary>
    public AsciiDocDocument SourceDocument { get; }

    /// <summary>Lossless document parsed from processed source.</summary>
    public AsciiDocDocument Document { get; }

    /// <summary>Expanded source passed to <see cref="Document"/>.</summary>
    public string ProcessedSource { get; }

    /// <summary>Effective attributes after preprocessing.</summary>
    public AsciiDocDocumentAttributes Attributes { get; }

    /// <summary>Processing diagnostics.</summary>
    public IReadOnlyList<AsciiDocProcessingDiagnostic> Diagnostics { get; }

    /// <summary>True when a processing error occurred.</summary>
    public bool HasErrors => Diagnostics.Any(static diagnostic => diagnostic.Severity == AsciiDocDiagnosticSeverity.Error);
}
