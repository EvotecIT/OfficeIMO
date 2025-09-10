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
}

