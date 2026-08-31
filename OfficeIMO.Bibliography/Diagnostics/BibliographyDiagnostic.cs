namespace OfficeIMO.Bibliography;

/// <summary>Severity of a bibliography diagnostic.</summary>
public enum BibliographyDiagnosticSeverity {
    /// <summary>Informational diagnostic.</summary>
    Information = 0,
    /// <summary>Recoverable condition that may affect semantics.</summary>
    Warning,
    /// <summary>Invalid or incomplete input.</summary>
    Error
}

/// <summary>A stable source-located parser or writer diagnostic.</summary>
public sealed class BibliographyDiagnostic {
    /// <summary>Initializes a diagnostic.</summary>
    public BibliographyDiagnostic(string code, BibliographyDiagnosticSeverity severity, string message, int offset = -1, int line = -1, int column = -1, string? itemKey = null, string? field = null) {
        Code = code ?? throw new ArgumentNullException(nameof(code));
        Severity = severity;
        Message = message ?? throw new ArgumentNullException(nameof(message));
        Offset = offset;
        Line = line;
        Column = column;
        ItemKey = itemKey;
        Field = field;
    }

    /// <summary>Stable machine-readable code.</summary>
    public string Code { get; }
    /// <summary>Diagnostic severity.</summary>
    public BibliographyDiagnosticSeverity Severity { get; }
    /// <summary>Human-readable description.</summary>
    public string Message { get; }
    /// <summary>Zero-based UTF-16 source offset, or -1.</summary>
    public int Offset { get; }
    /// <summary>One-based source line, or -1.</summary>
    public int Line { get; }
    /// <summary>One-based source column, or -1.</summary>
    public int Column { get; }
    /// <summary>Related citation key, when known.</summary>
    public string? ItemKey { get; }
    /// <summary>Related native field, when known.</summary>
    public string? Field { get; }

    /// <inheritdoc />
    public override string ToString() => $"{Code} {Severity}: {Message}";
}
