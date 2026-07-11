namespace OfficeIMO.Rtf;

/// <summary>Describes how one RTF feature was handled by a conversion boundary.</summary>
public sealed class RtfConversionDiagnostic {
    /// <summary>Initializes a conversion diagnostic.</summary>
    public RtfConversionDiagnostic(
        RtfConversionSeverity severity,
        string code,
        string message,
        RtfConversionAction action,
        string? sourcePath = null,
        string? feature = null,
        int count = 1,
        string? detail = null) {
        if (count <= 0) throw new ArgumentOutOfRangeException(nameof(count));
        Severity = severity;
        Code = code ?? throw new ArgumentNullException(nameof(code));
        Message = message ?? throw new ArgumentNullException(nameof(message));
        Action = action;
        SourcePath = sourcePath;
        Feature = feature;
        Count = count;
        Detail = detail;
    }

    /// <summary>Diagnostic severity.</summary>
    public RtfConversionSeverity Severity { get; }

    /// <summary>Stable machine-readable code.</summary>
    public string Code { get; }

    /// <summary>Human-readable description.</summary>
    public string Message { get; }

    /// <summary>Action taken for the source feature.</summary>
    public RtfConversionAction Action { get; }

    /// <summary>Optional logical path within the source document.</summary>
    public string? SourcePath { get; }

    /// <summary>Optional feature, element, destination, or control-word name.</summary>
    public string? Feature { get; }

    /// <summary>Number of equivalent occurrences represented by this entry.</summary>
    public int Count { get; }

    /// <summary>Optional adapter-specific details.</summary>
    public string? Detail { get; }
}
