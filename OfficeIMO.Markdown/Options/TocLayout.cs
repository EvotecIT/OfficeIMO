namespace OfficeIMO.Markdown;

/// <summary>
/// Visual layout variants for the HTML Table of Contents.
/// Affects container element/classes only (Markdown output is unchanged).
/// </summary>
public enum TocLayout {
    /// <summary>Render just a nested list (legacy behavior). No wrapper container is emitted.</summary>
    List = 0,
    /// <summary>Render inside a styled panel/card with optional title.</summary>
    Panel = 1,
    /// <summary>Render as a right-aligned sidebar (floats on wide screens, stacks on narrow).</summary>
    SidebarRight = 2,
    /// <summary>Render as a left-aligned sidebar (floats on wide screens, stacks on narrow).</summary>
    SidebarLeft = 3
}

