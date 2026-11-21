using OfficeIMO.Yaml;

namespace PSWriteOffice.Models.Yaml;

/// <summary>The outcome of Test-OfficeYaml -Detailed for one input.</summary>
public sealed class OfficeYamlTestResult {
    internal OfficeYamlTestResult(string? path, IReadOnlyList<OfficeDiagnostic> diagnostics, YamlException? error = null) {
        Path = path;
        Diagnostics = diagnostics;
        if (error != null) {
            OfficeDiagnostic failure = error.ToDiagnostic();
            Code = failure.Code;
            Message = failure.Message;
            Line = error.Line;
            Column = error.Column;
            Diagnostics = diagnostics.Append(failure).ToArray();
        }
    }

    /// <summary>The file tested, or null for text.</summary>
    public string? Path { get; }
    /// <summary>Whether the input loaded without an error.</summary>
    public bool IsValid => Code == null;
    /// <summary>The diagnostic code of the error, such as yaml-syntax-unexpected-token, or null when valid.</summary>
    public string? Code { get; }
    /// <summary>The error message with the offending source line, or null when valid.</summary>
    public string? Message { get; }
    /// <summary>The 1-based line of the error, or 0 when valid or unknown.</summary>
    public int Line { get; }
    /// <summary>The 1-based column of the error, or 0 when valid or unknown.</summary>
    public int Column { get; }
    /// <summary>
    /// Warnings and notes the engine recorded while loading, such as accepted lenient indentation or duplicate keys,
    /// followed by the error when the input does not load.
    /// </summary>
    public IReadOnlyList<OfficeDiagnostic> Diagnostics { get; }
}
