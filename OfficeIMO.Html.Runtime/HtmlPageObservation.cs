namespace OfficeIMO.Html.Runtime;

/// <summary>The page data included in an observation.</summary>
public enum HtmlPageObservationMode {
    /// <summary>Roles, accessible names, text, and form state.</summary>
    Semantic,
    /// <summary>Viewport geometry and visibility without semantic text.</summary>
    Visual,
    /// <summary>Semantic state and visual geometry.</summary>
    Combined
}

/// <summary>Bounds and filters for one revision-bound page observation.</summary>
public sealed class HtmlPageObservationRequest {
    /// <summary>Observation data to include.</summary>
    public HtmlPageObservationMode Mode { get; set; } = HtmlPageObservationMode.Combined;
    /// <summary>Maximum returned elements.</summary>
    public int MaxElements { get; set; } = 512;
    /// <summary>Maximum combined accessible-name and text characters.</summary>
    public int MaxTextCharacters { get; set; } = 256 * 1024;
    /// <summary>Includes elements hidden by markup or qualified layout.</summary>
    public bool IncludeHidden { get; set; }
    /// <summary>Returns only elements which qualify for a supported action.</summary>
    public bool ActionableOnly { get; set; }
    /// <summary>Requests a provider-owned screenshot artifact reference.</summary>
    public bool IncludeScreenshotReference { get; set; }

    /// <summary>Validates and returns a detached request within a provider's output-character budget.</summary>
    public HtmlPageObservationRequest Snapshot(int maximumOutputCharacters) {
        if (maximumOutputCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(maximumOutputCharacters));
        if (!Enum.IsDefined(Mode)) throw new ArgumentOutOfRangeException(nameof(Mode));
        if (MaxElements <= 0 || MaxElements > 10_000) throw new ArgumentOutOfRangeException(nameof(MaxElements));
        if (MaxTextCharacters <= 0 || MaxTextCharacters > maximumOutputCharacters)
            throw new ArgumentOutOfRangeException(nameof(MaxTextCharacters));
        return new HtmlPageObservationRequest {
            Mode = Mode,
            MaxElements = MaxElements,
            MaxTextCharacters = MaxTextCharacters,
            IncludeHidden = IncludeHidden,
            ActionableOnly = ActionableOnly,
            IncludeScreenshotReference = IncludeScreenshotReference
        };
    }
}

/// <summary>An opaque element identity valid only for the recorded page revision.</summary>
public sealed class HtmlObservedElementReference {
    /// <summary>Page which issued the reference.</summary>
    public string PageId { get; init; } = string.Empty;
    /// <summary>Page revision for which the reference is valid.</summary>
    public long Revision { get; init; }
    /// <summary>Element position in bounded document order.</summary>
    public int ElementIndex { get; init; }
    /// <summary>Element local name used to detect replacement.</summary>
    public string ElementName { get; init; } = string.Empty;
    /// <summary>Element id used to detect replacement.</summary>
    public string ElementId { get; init; } = string.Empty;
}

/// <summary>One semantic or visual element entry in document order.</summary>
public sealed class HtmlObservedElement {
    /// <summary>Revision-bound action target.</summary>
    public HtmlObservedElementReference Reference { get; init; } = null!;
    /// <summary>Element depth in the document tree.</summary>
    public int Depth { get; init; }
    /// <summary>Document-order index of the closest parent element.</summary>
    public int? ParentElementIndex { get; init; }
    /// <summary>Element local name.</summary>
    public string ElementName { get; init; } = string.Empty;
    /// <summary>Explicit or basic implicit accessibility role.</summary>
    public string Role { get; init; } = string.Empty;
    /// <summary>Bounded accessible name.</summary>
    public string AccessibleName { get; init; } = string.Empty;
    /// <summary>Bounded normalized descendant text.</summary>
    public string Text { get; init; } = string.Empty;
    /// <summary>Current non-password form-control value. Password values are always omitted.</summary>
    public string? Value { get; init; }
    /// <summary>Selected option values in document order.</summary>
    public IReadOnlyList<string> SelectedValues { get; init; } = Array.Empty<string>();
    /// <summary>Current UTF-16 selection start.</summary>
    public int? SelectionStart { get; init; }
    /// <summary>Current UTF-16 selection end.</summary>
    public int? SelectionEnd { get; init; }
    /// <summary>Current checkbox or radio state.</summary>
    public bool? IsChecked { get; init; }
    /// <summary>Whether form semantics disable the element.</summary>
    public bool IsDisabled { get; init; }
    /// <summary>Whether the element accepts the supported fill action.</summary>
    public bool IsEditable { get; init; }
    /// <summary>Whether the element owns page focus.</summary>
    public bool IsFocused { get; init; }
    /// <summary>Qualified layout visibility, when requested and available.</summary>
    public bool? IsVisible { get; init; }
    /// <summary>Qualified viewport intersection, when requested and available.</summary>
    public bool? IsInViewport { get; init; }
    /// <summary>Whether the element qualifies for a supported structured action.</summary>
    public bool IsActionable { get; init; }
    /// <summary>Bounding box in viewport CSS pixels, when requested and available.</summary>
    public HtmlRuntimeRect? BoundingBox { get; init; }
}

/// <summary>An optional provider-owned visual artifact referenced by an observation.</summary>
public sealed class HtmlObservationArtifactReference {
    /// <summary>Provider-owned artifact identity.</summary>
    public string Id { get; init; } = string.Empty;
    /// <summary>Artifact media type.</summary>
    public string MediaType { get; init; } = string.Empty;
    /// <summary>Lowercase hexadecimal SHA-256 digest.</summary>
    public string Sha256 { get; init; } = string.Empty;
    /// <summary>Artifact width in pixels.</summary>
    public int Width { get; init; }
    /// <summary>Artifact height in pixels.</summary>
    public int Height { get; init; }
}

/// <summary>A bounded page snapshot whose element references expire when its revision changes.</summary>
public sealed class HtmlPageObservation {
    /// <summary>Provider which produced the observation.</summary>
    public string ProviderId { get; init; } = string.Empty;
    /// <summary>Owning context identity.</summary>
    public string ContextId { get; init; } = string.Empty;
    /// <summary>Observed page identity.</summary>
    public string PageId { get; init; } = string.Empty;
    /// <summary>Revision binding all returned element references.</summary>
    public long Revision { get; init; }
    /// <summary>Current page URL.</summary>
    public Uri Url { get; init; } = new("https://officeimo.invalid/");
    /// <summary>Current document title.</summary>
    public string Title { get; init; } = string.Empty;
    /// <summary>Data mode used for this observation.</summary>
    public HtmlPageObservationMode Mode { get; init; }
    /// <summary>Viewport width in CSS pixels.</summary>
    public double ViewportWidth { get; init; }
    /// <summary>Viewport height in CSS pixels.</summary>
    public double ViewportHeight { get; init; }
    /// <summary>Horizontal document scroll offset.</summary>
    public double ScrollX { get; init; }
    /// <summary>Vertical document scroll offset.</summary>
    public double ScrollY { get; init; }
    /// <summary>Qualified document width.</summary>
    public double DocumentWidth { get; init; }
    /// <summary>Qualified document height.</summary>
    public double DocumentHeight { get; init; }
    /// <summary>Whether configured element or text bounds truncated the result.</summary>
    public bool IsTruncated { get; init; }
    /// <summary>Observed elements in document order after filtering.</summary>
    public IReadOnlyList<HtmlObservedElement> Elements { get; init; } = Array.Empty<HtmlObservedElement>();
    /// <summary>Optional provider-owned visual artifact references.</summary>
    public IReadOnlyList<HtmlObservationArtifactReference> Artifacts { get; init; } = Array.Empty<HtmlObservationArtifactReference>();
    /// <summary>Provider-neutral qualification or limitation notes.</summary>
    public IReadOnlyList<string> Diagnostics { get; init; } = Array.Empty<string>();
}
