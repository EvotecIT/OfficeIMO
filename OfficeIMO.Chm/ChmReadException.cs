namespace OfficeIMO.Chm;

/// <summary>A malformed, unsupported, or over-budget CHM input. No partial book is returned.</summary>
public sealed class ChmReadException : IOException {
    /// <summary>Creates a format failure with a stable diagnostic code.</summary>
    public ChmReadException(string code, string message, Exception? innerException = null) : base(message, innerException) { Code = code; }
    /// <summary>Stable machine-readable failure code.</summary>
    public string Code { get; }
}

/// <summary>A non-fatal qualification or navigation issue retained by a loaded book.</summary>
public sealed class ChmDiagnostic {
    internal ChmDiagnostic(string code, string message, string? path = null) { Code = code; Message = message; Path = path; }
    /// <summary>Stable machine-readable diagnostic code.</summary>
    public string Code { get; }
    /// <summary>Explanation of the affected behavior.</summary>
    public string Message { get; }
    /// <summary>Archive path or reference associated with the issue.</summary>
    public string? Path { get; }
}
