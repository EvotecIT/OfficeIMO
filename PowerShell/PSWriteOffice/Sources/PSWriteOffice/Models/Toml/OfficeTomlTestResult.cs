using OfficeIMO.Toml;

namespace PSWriteOffice.Models.Toml;

/// <summary>The outcome of Test-OfficeToml -Detailed for one input.</summary>
public sealed class OfficeTomlTestResult {
    internal OfficeTomlTestResult(string? path, TomlException? error = null) {
        Path = path;
        if (error != null) {
            OfficeDiagnostic failure = error.ToDiagnostic();
            Code = failure.Code;
            Message = failure.Message;
            Line = error.Line;
            Column = error.Column;
        }
    }

    /// <summary>The file tested, or null for text.</summary>
    public string? Path { get; }
    /// <summary>Whether the input loaded without an error.</summary>
    public bool IsValid => Code == null;
    /// <summary>The diagnostic code of the error, such as toml-load-duplicate-key, or null when valid.</summary>
    public string? Code { get; }
    /// <summary>The error message without its position, or null when valid.</summary>
    public string? Message { get; }
    /// <summary>The 1-based line of the error, or 0 when valid or unknown.</summary>
    public int Line { get; }
    /// <summary>The 1-based column of the error, or 0 when valid or unknown.</summary>
    public int Column { get; }
}
