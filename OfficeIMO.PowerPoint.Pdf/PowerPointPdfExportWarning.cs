namespace OfficeIMO.PowerPoint.Pdf;

/// <summary>
/// Describes PowerPoint content that the first-party PDF adapter skipped or simplified.
/// </summary>
public sealed class PowerPointPdfExportWarning {
    /// <summary>Creates a PowerPoint-to-PDF export warning.</summary>
    public PowerPointPdfExportWarning(int slideNumber, string code, string message) {
        SlideNumber = slideNumber;
        Code = code ?? throw new ArgumentNullException(nameof(code));
        Message = message ?? throw new ArgumentNullException(nameof(message));
    }

    /// <summary>One-based slide number where the warning originated.</summary>
    public int SlideNumber { get; }

    /// <summary>Stable warning code for callers that want to group export diagnostics.</summary>
    public string Code { get; }

    /// <summary>Human-readable warning message.</summary>
    public string Message { get; }
}
