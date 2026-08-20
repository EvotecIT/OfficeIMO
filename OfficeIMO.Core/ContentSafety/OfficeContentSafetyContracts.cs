using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OfficeIMO.ContentSafety;

/// <summary>Classifies how content can remain machine-readable while being concealed from an ordinary human view.</summary>
public enum OfficeContentConcealmentKind {
    /// <summary>A native format property explicitly hides the content.</summary>
    HiddenByProperty,
    /// <summary>An owning sheet, slide, layer, row, column, or similar container is hidden.</summary>
    HiddenContainer,
    /// <summary>The content uses a zero or near-zero font size.</summary>
    TinyText,
    /// <summary>The content or its owning geometry has no meaningful visible width or height.</summary>
    ZeroDimension,
    /// <summary>The text is fully or nearly transparent.</summary>
    TransparentText,
    /// <summary>The effective foreground and background have insufficient contrast.</summary>
    LowContrastText,
    /// <summary>The content is positioned outside the ordinary visible canvas.</summary>
    OffCanvas,
    /// <summary>The content is removed by an active clipping or overflow boundary.</summary>
    ClippedContent,
    /// <summary>The format uses a non-painting text rendering mode.</summary>
    InvisibleRenderingMode,
    /// <summary>A display format conceals a stored value while retaining it for machine access.</summary>
    HiddenDisplayValue,
    /// <summary>Content lives in notes, comments, metadata, alternative text, or another non-primary story.</summary>
    NonPrimaryContent,
    /// <summary>Text contains non-printing Unicode or malformed Unicode evidence.</summary>
    NonPrintingUnicode,
    /// <summary>Content uses an exact native concealment mechanism not represented by a more specific value.</summary>
    Other
}

/// <summary>Describes how a content-safety finding should be interpreted.</summary>
public enum OfficeContentSafetyRisk {
    /// <summary>The content is commonly legitimate but is reported so callers can preserve an exact audit trail.</summary>
    Informational,
    /// <summary>The content can be legitimate or adversarial and needs workflow context.</summary>
    ContextDependent,
    /// <summary>The content can materially influence machine processing while avoiding ordinary human review.</summary>
    PotentiallyDangerous
}

/// <summary>Describes whether a finding can be removed without deleting a wider container.</summary>
public enum OfficeContentCleanupCapability {
    /// <summary>The adapter cannot safely mutate this exact finding.</summary>
    ReportOnly,
    /// <summary>The adapter can remove the exact text payload while retaining its owner.</summary>
    RemoveText,
    /// <summary>The adapter can remove the exact owning element, run, cell, shape, or metadata value.</summary>
    RemoveElement
}

/// <summary>Bounds concealment and indirect prompt-injection inspection.</summary>
public sealed class OfficeContentSafetyOptions {
    /// <summary>Maximum encoded input bytes accepted before a parser or package loader runs. Defaults to 256 MiB.</summary>
    public long MaxInputBytes { get; set; } = 256L * 1024L * 1024L;
    /// <summary>Maximum ZIP package entries accepted during package preflight. Defaults to 16,384.</summary>
    public int MaxPackageEntries { get; set; } = 16 * 1024;
    /// <summary>Maximum declared uncompressed bytes accepted across a ZIP package. Defaults to 512 MiB.</summary>
    public long MaxExpandedPackageBytes { get; set; } = 512L * 1024L * 1024L;
    /// <summary>Maximum decoded characters inspected across the asset. Defaults to 16 million.</summary>
    public int MaxCharacters { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum findings returned. Defaults to 4,096.</summary>
    public int MaxFindings { get; set; } = 4096;
    /// <summary>Maximum characters retained in a finding preview. Defaults to 240.</summary>
    public int MaxPreviewCharacters { get; set; } = 240;
    /// <summary>Font sizes at or below this value are treated as effectively concealed. Defaults to two points.</summary>
    public double MaximumTinyFontSizePoints { get; set; } = 2D;
    /// <summary>Contrast ratios below this value are reported. Defaults to 1.25.</summary>
    public double MinimumVisibleContrastRatio { get; set; } = 1.25D;
    /// <summary>Whether notes, comments, metadata, alternative text, and other non-primary stories are reported.</summary>
    public bool IncludeNonPrimaryContent { get; set; } = true;
    /// <summary>Whether concealed text is checked for bounded instruction-like language.</summary>
    public bool DetectInstructionLikeText { get; set; } = true;
    /// <summary>Whether exact Unicode text-integrity evidence is included. Defaults to true.</summary>
    public bool IncludeTextIntegrityEvidence { get; set; } = true;

    internal void Validate() {
        if (MaxInputBytes <= 0 || MaxInputBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(MaxInputBytes));
        if (MaxPackageEntries <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPackageEntries));
        if (MaxExpandedPackageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxExpandedPackageBytes));
        if (MaxCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaxCharacters));
        if (MaxFindings <= 0) throw new ArgumentOutOfRangeException(nameof(MaxFindings));
        if (MaxPreviewCharacters <= 0 || MaxPreviewCharacters > 4096) throw new ArgumentOutOfRangeException(nameof(MaxPreviewCharacters));
        if (double.IsNaN(MaximumTinyFontSizePoints) || double.IsInfinity(MaximumTinyFontSizePoints) || MaximumTinyFontSizePoints < 0D || MaximumTinyFontSizePoints > 72D) {
            throw new ArgumentOutOfRangeException(nameof(MaximumTinyFontSizePoints));
        }
        if (double.IsNaN(MinimumVisibleContrastRatio) || double.IsInfinity(MinimumVisibleContrastRatio) || MinimumVisibleContrastRatio < 1D || MinimumVisibleContrastRatio > 21D) {
            throw new ArgumentOutOfRangeException(nameof(MinimumVisibleContrastRatio));
        }
    }
}

/// <summary>One exact concealed-content or machine-only-content finding.</summary>
public sealed class OfficeContentSafetyFinding {
    /// <summary>Creates an immutable content-safety finding.</summary>
    public OfficeContentSafetyFinding(
        string id,
        string format,
        OfficeContentConcealmentKind kind,
        OfficeContentSafetyRisk risk,
        string location,
        string evidence,
        string textPreview,
        int textLength,
        string contentHash,
        bool isInstructionLike,
        IReadOnlyList<string>? instructionSignals = null,
        OfficeContentCleanupCapability cleanupCapability = OfficeContentCleanupCapability.ReportOnly,
        int? sourceTextOffset = null,
        int? sourceTextLength = null) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("A stable finding id is required.", nameof(id));
        if (string.IsNullOrWhiteSpace(format)) throw new ArgumentException("A format name is required.", nameof(format));
        if (string.IsNullOrWhiteSpace(location)) throw new ArgumentException("A logical location is required.", nameof(location));
        if (string.IsNullOrWhiteSpace(evidence)) throw new ArgumentException("Exact concealment evidence is required.", nameof(evidence));
        if (textLength < 0) throw new ArgumentOutOfRangeException(nameof(textLength));
        if (sourceTextOffset.HasValue && sourceTextOffset.Value < 0) throw new ArgumentOutOfRangeException(nameof(sourceTextOffset));
        if (sourceTextLength.HasValue && sourceTextLength.Value <= 0) throw new ArgumentOutOfRangeException(nameof(sourceTextLength));
        if (string.IsNullOrWhiteSpace(contentHash)) throw new ArgumentException("A content hash is required.", nameof(contentHash));
        Id = id;
        Format = format;
        Kind = kind;
        Risk = risk;
        Location = location;
        Evidence = evidence;
        TextPreview = textPreview ?? string.Empty;
        TextLength = textLength;
        ContentHash = contentHash;
        IsInstructionLike = isInstructionLike;
        InstructionSignals = new List<string>(instructionSignals ?? Array.Empty<string>()).AsReadOnly();
        CleanupCapability = cleanupCapability;
        SourceTextOffset = sourceTextOffset;
        SourceTextLength = sourceTextLength;
    }

    /// <summary>Gets a deterministic id bound to the format, location, mechanism, and exact content hash.</summary>
    public string Id { get; }
    /// <summary>Gets the adapter format name.</summary>
    public string Format { get; }
    /// <summary>Gets the exact concealment mechanism.</summary>
    public OfficeContentConcealmentKind Kind { get; }
    /// <summary>Gets the interpretation risk.</summary>
    public OfficeContentSafetyRisk Risk { get; }
    /// <summary>Gets the format-native logical location.</summary>
    public string Location { get; }
    /// <summary>Gets bounded, adapter-produced evidence explaining why the content is concealed.</summary>
    public string Evidence { get; }
    /// <summary>Gets a bounded, control-safe preview. This is never used as a mutation locator.</summary>
    public string TextPreview { get; }
    /// <summary>Gets the original UTF-16 text length.</summary>
    public int TextLength { get; }
    /// <summary>Gets the SHA-256 hash of the exact text payload.</summary>
    public string ContentHash { get; }
    /// <summary>Gets whether bounded heuristics found instruction-like language in concealed content.</summary>
    public bool IsInstructionLike { get; }
    /// <summary>Gets the exact heuristic signal identifiers that were observed.</summary>
    public IReadOnlyList<string> InstructionSignals { get; }
    /// <summary>Gets whether the owning adapter can safely remove this exact finding.</summary>
    public OfficeContentCleanupCapability CleanupCapability { get; }
    /// <summary>Gets the exact UTF-16 offset within the reported native text surface when this is a Unicode finding.</summary>
    public int? SourceTextOffset { get; }
    /// <summary>Gets the exact UTF-16 length within the reported native text surface when this is a Unicode finding.</summary>
    public int? SourceTextLength { get; }
}

/// <summary>Immutable content-safety evidence for one asset.</summary>
public sealed class OfficeContentSafetyReport {
    /// <summary>Creates a report from format findings and exact Unicode evidence.</summary>
    public OfficeContentSafetyReport(
        string format,
        IReadOnlyList<OfficeContentSafetyFinding> findings,
        IReadOnlyList<OfficeIMO.Provenance.OfficeTextIntegrityFinding>? textIntegrityFindings = null,
        IReadOnlyList<string>? diagnostics = null) {
        if (string.IsNullOrWhiteSpace(format)) throw new ArgumentException("A format name is required.", nameof(format));
        Format = format;
        Findings = new List<OfficeContentSafetyFinding>(findings ?? throw new ArgumentNullException(nameof(findings))).AsReadOnly();
        TextIntegrityFindings = new List<OfficeIMO.Provenance.OfficeTextIntegrityFinding>(textIntegrityFindings ?? Array.Empty<OfficeIMO.Provenance.OfficeTextIntegrityFinding>()).AsReadOnly();
        Diagnostics = new List<string>(diagnostics ?? Array.Empty<string>()).AsReadOnly();
    }

    /// <summary>Gets the inspected format name.</summary>
    public string Format { get; }
    /// <summary>Gets concealment findings in deterministic source order.</summary>
    public IReadOnlyList<OfficeContentSafetyFinding> Findings { get; }
    /// <summary>Gets exact Unicode evidence collected from inspected text surfaces.</summary>
    public IReadOnlyList<OfficeIMO.Provenance.OfficeTextIntegrityFinding> TextIntegrityFindings { get; }
    /// <summary>Gets bounded diagnostics for unsupported or ambiguous format constructs.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
    /// <summary>Gets whether concealed instruction-like content or dangerous Unicode evidence was found.</summary>
    public bool HasPotentiallyDangerousContent =>
        Findings.Any(item => item.Risk == OfficeContentSafetyRisk.PotentiallyDangerous) ||
        TextIntegrityFindings.Any(item => item.Risk == OfficeIMO.Provenance.OfficeTextIntegrityRisk.PotentiallyDangerous);
    /// <summary>Gets whether any machine-readable content is concealed or outside the primary human-visible story.</summary>
    public bool HasConcealedContent => Findings.Any(item => item.Kind != OfficeContentConcealmentKind.NonPrintingUnicode);
}

/// <summary>Selects exact findings for an explicit cleanup operation.</summary>
public sealed class OfficeContentCleanupSelection {
    /// <summary>Creates a selection from stable finding ids.</summary>
    public OfficeContentCleanupSelection(IEnumerable<string> findingIds) {
        if (findingIds == null) throw new ArgumentNullException(nameof(findingIds));
        string[] ids = findingIds.Select(item => item?.Trim() ?? string.Empty).ToArray();
        if (ids.Any(string.IsNullOrWhiteSpace)) throw new ArgumentException("Finding ids cannot be empty.", nameof(findingIds));
        if (ids.Length != ids.Distinct(StringComparer.Ordinal).Count()) throw new ArgumentException("Finding ids must be unique.", nameof(findingIds));
        FindingIds = Array.AsReadOnly(ids);
    }

    /// <summary>Gets the selected stable finding ids.</summary>
    public IReadOnlyList<string> FindingIds { get; }
}

/// <summary>Controls explicit artifact cleanup and signature invalidation policy.</summary>
public sealed class OfficeContentCleanupOptions {
    /// <summary>Inspection thresholds used before and after cleanup.</summary>
    public OfficeContentSafetyOptions Inspection { get; set; } = new OfficeContentSafetyOptions();
    /// <summary>How package mutation handles existing digital signatures. Defaults to fail closed.</summary>
    public OfficeIMO.OfficeSignatureMutationPolicy SignatureMutationPolicy { get; set; } = OfficeIMO.OfficeSignatureMutationPolicy.BlockSave;

    internal void Validate() {
        if (Inspection == null) throw new ArgumentNullException(nameof(Inspection));
        Inspection.Validate();
        if (!Enum.IsDefined(typeof(OfficeIMO.OfficeSignatureMutationPolicy), SignatureMutationPolicy)) {
            throw new ArgumentOutOfRangeException(nameof(SignatureMutationPolicy));
        }
    }
}

/// <summary>One exact cleanup change.</summary>
public sealed class OfficeContentCleanupChange {
    /// <summary>Creates a cleanup change.</summary>
    public OfficeContentCleanupChange(string findingId, string location, OfficeContentCleanupCapability capability) {
        if (string.IsNullOrWhiteSpace(findingId)) throw new ArgumentException("A finding id is required.", nameof(findingId));
        if (string.IsNullOrWhiteSpace(location)) throw new ArgumentException("A location is required.", nameof(location));
        FindingId = findingId;
        Location = location;
        Capability = capability;
    }
    /// <summary>Gets the removed finding id.</summary>
    public string FindingId { get; }
    /// <summary>Gets the format-native location that changed.</summary>
    public string Location { get; }
    /// <summary>Gets the exact cleanup operation applied.</summary>
    public OfficeContentCleanupCapability Capability { get; }
}

/// <summary>Result of an explicit, selection-based content cleanup.</summary>
public sealed class OfficeContentCleanupResult {
    /// <summary>Creates a cleanup result.</summary>
    public OfficeContentCleanupResult(
        byte[] output,
        OfficeContentSafetyReport before,
        OfficeContentSafetyReport after,
        IReadOnlyList<OfficeContentCleanupChange> changes) {
        Output = (byte[])(output ?? throw new ArgumentNullException(nameof(output))).Clone();
        Before = before ?? throw new ArgumentNullException(nameof(before));
        After = after ?? throw new ArgumentNullException(nameof(after));
        Changes = new List<OfficeContentCleanupChange>(changes ?? throw new ArgumentNullException(nameof(changes))).AsReadOnly();
    }
    /// <summary>Gets a defensive copy of the cleaned artifact.</summary>
    public byte[] Output { get; }
    /// <summary>Gets the evidence snapshot before cleanup.</summary>
    public OfficeContentSafetyReport Before { get; }
    /// <summary>Gets the evidence snapshot after cleanup.</summary>
    public OfficeContentSafetyReport After { get; }
    /// <summary>Gets exact changes in selection order.</summary>
    public IReadOnlyList<OfficeContentCleanupChange> Changes { get; }
    /// <summary>Gets whether the output differs from the inspected input.</summary>
    public bool Changed => Changes.Count > 0;
}
