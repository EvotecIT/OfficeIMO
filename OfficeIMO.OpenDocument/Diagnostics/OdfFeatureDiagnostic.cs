namespace OfficeIMO.OpenDocument;

/// <summary>Diagnostic produced while inspecting package capabilities.</summary>
public sealed class OdfFeatureDiagnostic {
    /// <summary>Creates an inspection diagnostic.</summary>
    public OdfFeatureDiagnostic(string code, string partPath, string message) {
        if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("Diagnostic code cannot be empty.", nameof(code));
        if (string.IsNullOrWhiteSpace(partPath)) throw new ArgumentException("Part path cannot be empty.", nameof(partPath));
        Code = code;
        PartPath = partPath;
        Message = message ?? throw new ArgumentNullException(nameof(message));
    }

    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }
    /// <summary>Package part that could not be inspected.</summary>
    public string PartPath { get; }
    /// <summary>Human-readable diagnostic.</summary>
    public string Message { get; }
}
