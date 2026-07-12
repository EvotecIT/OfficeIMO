namespace OfficeIMO.Pdf;

/// <summary>One stable reason supporting or weakening an inferred PDF artifact.</summary>
public sealed class PdfInferenceEvidence {
    /// <summary>Creates diagnostic evidence with a normalized contribution.</summary>
    public PdfInferenceEvidence(string code, string message, double contribution) {
        Guard.NotNullOrWhiteSpace(code, nameof(code));
        Guard.NotNullOrWhiteSpace(message, nameof(message));
        if (double.IsNaN(contribution) || double.IsInfinity(contribution) || contribution < -1D || contribution > 1D) throw new ArgumentOutOfRangeException(nameof(contribution));
        Code = code; Message = message; Contribution = contribution;
    }
    /// <summary>Stable machine-readable reason code.</summary>
    public string Code { get; }
    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }
    /// <summary>Normalized contribution from -1 (weakening) to 1 (supporting).</summary>
    public double Contribution { get; }
}

/// <summary>Evidence for one region's inferred position in page reading order.</summary>
public sealed class PdfReadingOrderEvidence {
    internal PdfReadingOrderEvidence(int index, PdfUnderstandingRegion region, double confidence, IReadOnlyList<PdfInferenceEvidence> evidence) {
        Index = index; Region = region; Confidence = confidence; Evidence = evidence;
    }
    /// <summary>Zero-based inferred reading position.</summary>
    public int Index { get; }
    /// <summary>Region assigned to this reading position.</summary>
    public PdfUnderstandingRegion Region { get; }
    /// <summary>Normalized confidence from 0 to 1.</summary>
    public double Confidence { get; }
    /// <summary>Diagnostic evidence supporting the order.</summary>
    public IReadOnlyList<PdfInferenceEvidence> Evidence { get; }
}

internal static class PdfInference {
    internal static double Clamp(double value) => value <= 0D ? 0D : value >= 1D ? 1D : value;
    internal static IReadOnlyList<PdfInferenceEvidence> Snapshot(IEnumerable<PdfInferenceEvidence>? evidence) => evidence?.ToArray() ?? Array.Empty<PdfInferenceEvidence>();
}
