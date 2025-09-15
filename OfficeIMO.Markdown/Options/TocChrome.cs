namespace OfficeIMO.Markdown;

/// <summary>
/// Visual chrome level for the TOC container.
/// </summary>
public enum TocChrome {
    /// <summary>Use style defaults (panel for Panel layout; plain for sidebars).</summary>
    Default = 0,
    /// <summary>No background or border; text-only ("ghost").</summary>
    None = 1,
    /// <summary>Border only; transparent background.</summary>
    Outline = 2,
    /// <summary>Card-style panel; background + border.</summary>
    Panel = 3
}

