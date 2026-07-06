namespace OfficeIMO.Pdf;

/// <summary>
/// Describes one preservation mismatch found after rewriting or manipulating a PDF.
/// </summary>
public sealed class PdfRewritePreservationIssue {
    internal PdfRewritePreservationIssue(string feature, string expected, string actual, string message) {
        Feature = feature;
        Expected = expected;
        Actual = actual;
        Message = message;
    }

    /// <summary>Stable feature name, for example PageCount, Metadata.Title, or LinkAnnotations.</summary>
    public string Feature { get; }

    /// <summary>Expected value captured from the original PDF or declared options.</summary>
    public string Expected { get; }

    /// <summary>Actual value observed in the rewritten PDF.</summary>
    public string Actual { get; }

    /// <summary>Human-readable preservation diagnostic.</summary>
    public string Message { get; }
}
