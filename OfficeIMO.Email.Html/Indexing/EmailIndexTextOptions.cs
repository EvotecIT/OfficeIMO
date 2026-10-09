namespace OfficeIMO.Email;

/// <summary>Controls a bounded indexing projection. Quote and signature exclusions are opt-in heuristics.</summary>
public sealed class EmailIndexTextOptions {
    /// <summary>Inspect the selected body and include advisory evidence in the result.</summary>
    public bool InspectContentSafety { get; set; }
    /// <summary>Controls concealed HTML in indexing text. Exclusion also enables inspection.</summary>
    public EmailConcealedTextPolicy ConcealedTextPolicy { get; set; }
    /// <summary>Maximum source characters accepted before parsing; oversized input is rejected.</summary>
    public int MaxSourceChars { get; set; } = 2 * 1024 * 1024;
    /// <summary>Separate maximum for generated HTML after text encoding or RTF conversion.</summary>
    public int MaxProjectionChars { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum projected UTF-16 characters; clipping preserves Unicode scalar boundaries.</summary>
    public int MaxTextChars { get; set; } = 256 * 1024;
    /// <summary>Prefer the supplied plain-text alternative when available.</summary>
    public bool PreferPlainText { get; set; } = true;
    /// <summary>Exclude recognized quoted regions from SelectedText; FullText remains available.</summary>
    public bool ExcludeQuotes { get; set; }
    /// <summary>Exclude recognized signature regions from SelectedText; FullText remains available.</summary>
    public bool ExcludeSignatures { get; set; }
}

/// <summary>Evidence for a text region; unclassified text is not a proof of authorship.</summary>
public enum EmailIndexRegionKind {
    /// <summary>No supported quote or signature marker was recognized.</summary>
    Unclassified,
    /// <summary>A supported quote marker was recognized.</summary>
    Quoted,
    /// <summary>A supported signature marker was recognized.</summary>
    Signature
}

/// <summary>A contiguous region in FullText, with UTF-16 offsets and an inspectable classification reason.</summary>
public sealed class EmailIndexTextRegion {
    internal EmailIndexTextRegion(int start, int length, EmailIndexRegionKind kind, string reason) {
        Start = start; Length = length; Kind = kind; Reason = reason;
    }
    /// <summary>Zero-based UTF-16 offset in FullText.</summary>
    public int Start { get; }
    /// <summary>UTF-16 length in FullText.</summary>
    public int Length { get; }
    /// <summary>Recognized region category.</summary>
    public EmailIndexRegionKind Kind { get; }
    /// <summary>Marker responsible for classification, or unclassified.</summary>
    public string Reason { get; }
}

/// <summary>Markup-free text and the evidence used for optional exclusions.</summary>
public sealed class EmailIndexTextResult {
    /// <summary>Selected-body evidence without private text; null when inspection was not requested.</summary>
    public EmailBodyContentSafetyReport? ContentSafety { get; internal set; }
    internal EmailIndexTextResult(string fullText, string selectedText, EmailBodySourceKind sourceKind,
        IReadOnlyList<EmailIndexTextRegion> regions, bool truncated, IReadOnlyList<EmailDiagnostic> diagnostics) {
        FullText = fullText; SelectedText = selectedText; SourceKind = sourceKind;
        Regions = regions; Truncated = truncated; Diagnostics = diagnostics;
    }
    /// <summary>Complete projected text within the requested character bound, before quote/signature exclusions.</summary>
    public string FullText { get; }
    /// <summary>Text after explicitly requested region exclusions.</summary>
    public string SelectedText { get; }
    /// <summary>Selected body representation.</summary>
    public EmailBodySourceKind SourceKind { get; }
    /// <summary>Ordered, non-overlapping regions covering FullText.</summary>
    public IReadOnlyList<EmailIndexTextRegion> Regions { get; }
    /// <summary>Whether the text character limit omitted content.</summary>
    public bool Truncated { get; }
    /// <summary>Body policy and truncation evidence.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
}
