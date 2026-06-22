namespace OfficeIMO.Pdf;

/// <summary>Filters for removing PDF annotations during a safe full rewrite.</summary>
public sealed class PdfAnnotationRemovalOptions {
    /// <summary>Specific annotation object number to remove. When omitted, other filters decide the match.</summary>
    public int? ObjectNumber { get; set; }

    /// <summary>One-based page number whose annotations should be considered. When omitted, all pages are considered.</summary>
    public int? PageNumber { get; set; }

    /// <summary>Annotation subtype, for example Text, Link, or Widget. When omitted, all subtypes are considered.</summary>
    public string? Subtype { get; set; }

    /// <summary>Remove popup annotations linked from matching annotations through /Popup.</summary>
    public bool RemoveMatchingPopups { get; set; } = true;
}
