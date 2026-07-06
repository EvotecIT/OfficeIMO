namespace OfficeIMO.Pdf;

/// <summary>
/// Describes one missing or failed item in a PDF conversion proof snapshot.
/// </summary>
public sealed class PdfConversionProofIssue {
    internal PdfConversionProofIssue(string feature, string expected, string actual) {
        Feature = string.IsNullOrWhiteSpace(feature) ? "ConversionProof" : feature;
        Expected = expected ?? string.Empty;
        Actual = actual ?? string.Empty;
    }

    /// <summary>Proof feature that failed, such as text marker or warning code evidence.</summary>
    public string Feature { get; }

    /// <summary>Expected proof value.</summary>
    public string Expected { get; }

    /// <summary>Actual proof value observed in the converted PDF or conversion report.</summary>
    public string Actual { get; }

    /// <summary>Human-readable issue summary.</summary>
    public string Message => Feature + " expected " + Expected + " but found " + Actual + ".";
}
