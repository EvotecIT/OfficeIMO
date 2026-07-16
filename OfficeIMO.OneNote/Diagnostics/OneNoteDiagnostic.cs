namespace OfficeIMO.OneNote;

/// <summary>
/// Severity of a OneNote parsing or writing diagnostic.
/// </summary>
public enum OneNoteDiagnosticSeverity {
    /// <summary>Additional format information.</summary>
    Information = 0,
    /// <summary>A recoverable compatibility or fidelity concern.</summary>
    Warning = 1,
    /// <summary>An error that prevents the requested operation.</summary>
    Error = 2
}

/// <summary>
/// A structured diagnostic emitted while reading or writing OneNote data.
/// </summary>
public sealed class OneNoteDiagnostic {
    /// <summary>Stable diagnostic identifier.</summary>
    public string Code { get; set; } = string.Empty;

    /// <summary>Diagnostic severity.</summary>
    public OneNoteDiagnosticSeverity Severity { get; set; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; set; } = string.Empty;

    /// <summary>Optional zero-based byte offset associated with the diagnostic.</summary>
    public long? Offset { get; set; }

    /// <summary>Optional source path associated with the diagnostic.</summary>
    public string? SourcePath { get; set; }
}

/// <summary>
/// Exception raised when OneNote binary data violates a required format or safety invariant.
/// </summary>
public sealed class OneNoteFormatException : IOException {
    /// <summary>Creates a format exception.</summary>
    public OneNoteFormatException(string code, string message, long? offset = null, Exception? innerException = null)
        : base(message, innerException) {
        Code = string.IsNullOrWhiteSpace(code) ? "ONENOTE_FORMAT" : code;
        Offset = offset;
    }

    /// <summary>Stable error identifier.</summary>
    public string Code { get; }

    /// <summary>Optional zero-based byte offset where the failure was detected.</summary>
    public long? Offset { get; }
}
