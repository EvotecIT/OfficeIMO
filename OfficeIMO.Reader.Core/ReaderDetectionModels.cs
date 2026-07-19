using System;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

/// <summary>
/// Controls how Reader combines extension and content evidence.
/// </summary>
public enum ReaderDetectionMode {
    /// <summary>Use only the source name or path extension.</summary>
    ExtensionOnly = 0,

    /// <summary>Inspect content only when the extension does not identify a supported kind.</summary>
    ContentWhenUnknown = 1,

    /// <summary>Inspect content and prefer medium- or high-confidence content evidence.</summary>
    PreferContent = 2
}

/// <summary>
/// Confidence assigned to a detection result.
/// </summary>
public enum ReaderDetectionConfidence {
    /// <summary>No usable evidence was found.</summary>
    None = 0,

    /// <summary>Weak or ambiguous evidence, such as mostly textual bytes.</summary>
    Low = 1,

    /// <summary>Useful structural evidence, such as an extension or leading text token.</summary>
    Medium = 2,

    /// <summary>Strong signature or container evidence.</summary>
    High = 3
}

/// <summary>
/// Bounds and policy for content detection.
/// </summary>
public sealed class ReaderDetectionOptions {
    /// <summary>
    /// Detection policy. Standalone detection defaults to <see cref="ReaderDetectionMode.PreferContent"/>.
    /// </summary>
    public ReaderDetectionMode Mode { get; set; } = ReaderDetectionMode.PreferContent;

    /// <summary>
    /// Maximum prefix bytes inspected for signatures and text markers. Default: 64 KiB.
    /// </summary>
    public int MaxProbeBytes { get; set; } = 64 * 1024;

    /// <summary>
    /// Maximum archive entries inspected while classifying ZIP-based formats. Default: 512.
    /// </summary>
    public int MaxContainerEntries { get; set; } = 512;

    /// <summary>
    /// When true, inspect seekable ZIP containers for Office, Visio, and EPUB package markers. Default: true.
    /// </summary>
    public bool InspectContainers { get; set; } = true;
}

/// <summary>
/// Structured extension and content detection result.
/// </summary>
public sealed class ReaderDetectionResult {
    /// <summary>Source name or path used for extension evidence.</summary>
    public string SourceName { get; set; } = string.Empty;

    /// <summary>Normalized source extension, or an empty string.</summary>
    public string Extension { get; set; } = string.Empty;

    /// <summary>Kind inferred from the extension and active handler registry.</summary>
    public ReaderInputKind ExtensionKind { get; set; } = ReaderInputKind.Unknown;

    /// <summary>Kind inferred from content.</summary>
    public ReaderInputKind ContentKind { get; set; } = ReaderInputKind.Unknown;

    /// <summary>Confidence in extension evidence.</summary>
    public ReaderDetectionConfidence ExtensionConfidence { get; set; }

    /// <summary>Confidence in content evidence.</summary>
    public ReaderDetectionConfidence ContentConfidence { get; set; }

    /// <summary>Effective kind selected by the requested detection mode.</summary>
    public ReaderInputKind Kind { get; set; } = ReaderInputKind.Unknown;

    /// <summary>Confidence in the effective kind.</summary>
    public ReaderDetectionConfidence Confidence { get; set; }

    /// <summary>Detected media type when known.</summary>
    public string? MediaType { get; set; }

    /// <summary>True when content bytes were inspected.</summary>
    public bool ContentInspected { get; set; }

    /// <summary>True when a ZIP container was structurally inspected.</summary>
    public bool ContainerInspected { get; set; }

    /// <summary>Number of prefix bytes inspected.</summary>
    public int InspectedBytes { get; set; }

    /// <summary>Stable evidence tokens explaining the result.</summary>
    public IReadOnlyList<string> Evidence { get; set; } = Array.Empty<string>();

    /// <summary>
    /// True when medium- or high-confidence content evidence disagrees with a known extension kind.
    /// </summary>
    public bool IsMismatch =>
        ExtensionKind != ReaderInputKind.Unknown &&
        ContentKind != ReaderInputKind.Unknown &&
        ExtensionKind != ContentKind &&
        ContentInspected &&
        ContentConfidence >= ReaderDetectionConfidence.Medium;
}
