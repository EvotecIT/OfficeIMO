namespace OfficeIMO.Pdf;

/// <summary>Filters for removing PDF annotations during a safe full rewrite.</summary>
public sealed class PdfAnnotationRemovalOptions {
    /// <summary>Preferred mutation mode. Automatic uses a full rewrite when safe and append-only when required.</summary>
    public PdfMutationExecutionPreference ExecutionPreference { get; set; } = PdfMutationExecutionPreference.Automatic;

    /// <summary>Specific annotation object number to remove. When omitted, other filters decide the match.</summary>
    public int? ObjectNumber { get; set; }

    /// <summary>One-based page number whose annotations should be considered. When omitted, all pages are considered.</summary>
    public int? PageNumber { get; set; }

    /// <summary>Annotation subtype, for example Text, Link, or Widget. When omitted, all subtypes are considered.</summary>
    public string? Subtype { get; set; }

    /// <summary>Remove popup annotations linked from matching annotations through /Popup.</summary>
    public bool RemoveMatchingPopups { get; set; } = true;

    /// <summary>
    /// Allows append-only removal even though prior revisions retain the removed annotation bytes.
    /// Disabled by default because append-only output is not a sanitization boundary.
    /// </summary>
    public bool AllowResidualDataInAppendOnly { get; set; }
}
