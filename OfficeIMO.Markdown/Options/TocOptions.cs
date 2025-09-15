namespace OfficeIMO.Markdown;

/// <summary>
/// Options controlling TOC generation.
/// </summary>
public sealed class TocOptions {
    /// <summary>Minimum heading level to include (default 1).</summary>
    public int MinLevel { get; set; } = 1;
    /// <summary>Maximum heading level to include (default 3).</summary>
    public int MaxLevel { get; set; } = 3;
    /// <summary>When true, generates an ordered list; otherwise unordered (default).</summary>
    public bool Ordered { get; set; } = false;
    /// <summary>Include a title heading above the TOC (default true).</summary>
    public bool IncludeTitle { get; set; } = true;
    /// <summary>Title text (default "Table of Contents").</summary>
    public string Title { get; set; } = "Table of Contents";
    /// <summary>Heading level for the title (default 2).</summary>
    public int TitleLevel { get; set; } = 2;
    /// <summary>Limits TOC scope. Document = all headings; PreviousHeading = only headings under the nearest preceding heading; HeadingTitle = headings under the named heading.</summary>
    public TocScope Scope { get; set; } = TocScope.Document;
    /// <summary>For Scope=HeadingTitle, the heading text to scope under (case-insensitive).</summary>
    public string? ScopeHeadingTitle { get; set; }
    /// <summary>Render the TOC inside a collapsible HTML &lt;details&gt; element (HTML output only).</summary>
    public bool Collapsible { get; set; } = false;
    /// <summary>When <see cref="Collapsible"/> is true, render the TOC collapsed by default.</summary>
    public bool Collapsed { get; set; } = false;

    /// <summary>Visual layout variant for HTML rendering. Default <see cref="TocLayout.List"/> (legacy plain list).</summary>
    public TocLayout Layout { get; set; } = TocLayout.List;

    /// <summary>Enable ScrollSpy behavior (highlight active heading link while scrolling). HTML output only.</summary>
    public bool ScrollSpy { get; set; } = false;

    /// <summary>When true, the TOC container becomes sticky (position: sticky) with a small top offset.</summary>
    public bool Sticky { get; set; } = false;

    /// <summary>Optional ARIA label for the navigation container. Default: "Table of Contents".</summary>
    public string? AriaLabel { get; set; } = "Table of Contents";

    /// <summary>Optional sidebar width in pixels for SidebarLeft/SidebarRight layouts. Default 260.</summary>
    public int? WidthPx { get; set; }

    /// <summary>Visual chrome for the TOC container: None, Outline, Panel, or Default.</summary>
    public TocChrome Chrome { get; set; } = TocChrome.Default;

    /// <summary>When true, hides the sidebar TOC completely on narrow screens (width up to 1000px).</summary>
    public bool HideOnNarrow { get; set; } = false;

    /// <summary>
    /// When true, ensure the top-level (H1) is included even if a deeper MinLevel was specified.
    /// Helps avoid unreadable TOCs that start at H2/H3 without their parent context.
    /// </summary>
    public bool RequireTopLevel { get; set; } = true;

    /// <summary>
    /// When true (default), the visual indentation of TOC items is normalized so that the minimum included
    /// heading level aligns as the root. When false, indentation reflects absolute heading levels (H1 root).
    /// </summary>
    public bool NormalizeToMinLevel { get; set; } = true;
}
