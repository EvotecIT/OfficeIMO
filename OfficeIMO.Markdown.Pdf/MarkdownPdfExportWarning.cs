namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// Describes a Markdown to PDF export feature that was skipped or simplified.
/// </summary>
public sealed class MarkdownPdfExportWarning {
    /// <summary>Stable warning code for callers and wrappers.</summary>
    public string Code { get; }

    /// <summary>Markdown source area, block type, or path associated with the warning.</summary>
    public string Source { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Creates a Markdown PDF export warning.</summary>
    public MarkdownPdfExportWarning(string code, string source, string message) {
        Code = string.IsNullOrWhiteSpace(code) ? "MarkdownPdfWarning" : code;
        Source = source ?? string.Empty;
        Message = message ?? string.Empty;
    }

    /// <inheritdoc />
    public override string ToString() {
        return string.IsNullOrWhiteSpace(Source)
            ? Code + ": " + Message
            : Code + " [" + Source + "]: " + Message;
    }
}
