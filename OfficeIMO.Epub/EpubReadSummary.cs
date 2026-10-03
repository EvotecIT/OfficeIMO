namespace OfficeIMO.Epub;

/// <summary>
/// Describes completeness of chapter extraction under the requested selection policy.
/// This is not a conformance, rendering-fidelity, or resource-retention guarantee.
/// </summary>
public sealed class EpubReadSummary {
    /// <summary>Whether chapters were selected from a declared OPF spine.</summary>
    public bool IsSpineBased { get; internal set; }

    /// <summary>Whether recovery scanning supplied chapters without a usable spine.</summary>
    public bool UsedFallbackScan { get; internal set; }

    /// <summary>Selected reading positions, including missing or unsupported items.</summary>
    public int RequestedChapterCount { get; internal set; }

    /// <summary>Reading positions successfully emitted as chapters.</summary>
    public int ExtractedChapterCount { get; internal set; }

    /// <summary>Selected positions that could not be emitted, including configured limits.</summary>
    public int SkippedChapterCount => RequestedChapterCount - ExtractedChapterCount;

    /// <summary>
    /// Whether every selected spine position was emitted. Policy-excluded non-linear
    /// positions are not requested; recovery scanning never establishes completeness.
    /// </summary>
    public bool IsComplete => IsSpineBased && !UsedFallbackScan && SkippedChapterCount == 0;
}
