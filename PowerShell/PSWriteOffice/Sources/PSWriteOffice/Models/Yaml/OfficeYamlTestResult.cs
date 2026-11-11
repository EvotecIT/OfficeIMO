using OfficeIMO.Yaml;

namespace PSWriteOffice.Models.Yaml;

/// <summary>The outcome of Test-OfficeYaml -Detailed for one input.</summary>
/// <remarks>
/// A failure carries the engine's <see cref="YamlException"/> fields as they are. When the engine exposes its
/// YamlException-to-OfficeDiagnostic helper, the failure becomes one more entry in <see cref="Diagnostics"/>.
/// </remarks>
public sealed class OfficeYamlTestResult {
    internal OfficeYamlTestResult(string? path, YamlException? error, IReadOnlyList<OfficeDiagnostic> diagnostics) {
        Path = path;
        Diagnostics = diagnostics;
        if (error != null) {
            Code = error.Code;
            Message = error.Message;
            Line = error.Line;
            Column = error.Column;
        }
    }

    /// <summary>The file tested, or null for text.</summary>
    public string? Path { get; }
    /// <summary>Whether the input loaded without an error.</summary>
    public bool IsValid => Code == null;
    /// <summary>The engine's stable error code, such as Yaml.Syntax.UnexpectedToken, or null when valid.</summary>
    public string? Code { get; }
    /// <summary>The error message, or null when valid.</summary>
    public string? Message { get; }
    /// <summary>The 1-based line of the error, or 0 when valid or unknown.</summary>
    public int Line { get; }
    /// <summary>The 1-based column of the error, or 0 when valid or unknown.</summary>
    public int Column { get; }
    /// <summary>Warnings and notes the engine recorded while loading, such as accepted lenient indentation or duplicate keys.</summary>
    public IReadOnlyList<OfficeDiagnostic> Diagnostics { get; }
}
