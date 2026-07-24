using OfficeIMO.PowerPoint;

namespace OfficeIMO.Markup.PowerPoint;

/// <summary>Controls conversion of an Office Markup document to a PowerPoint presentation.</summary>
public sealed class MarkupToPowerPointOptions {
    /// <summary>Base directory used to resolve relative resource paths.</summary>
    public string? BaseDirectory { get; set; }
    /// <summary>Slide width in inches.</summary>
    public double SlideWidthInches { get; set; } = 10.0;
    /// <summary>Slide height in inches.</summary>
    public double SlideHeightInches { get; set; } = 5.625;
    /// <summary>Whether unsupported blocks should be preserved as visible text.</summary>
    public bool IncludeUnsupportedBlocksAsText { get; set; } = true;
    /// <summary>Whether image paths outside <see cref="BaseDirectory"/> may be read.</summary>
    public bool AllowExternalImagePaths { get; set; }
    /// <summary>Whether Mermaid diagram blocks should be rendered.</summary>
    public bool RenderMermaidDiagrams { get; set; } = true;
    /// <summary>Optional Mermaid renderer executable path.</summary>
    public string? MermaidRendererPath { get; set; }
    /// <summary>Optional directory for temporary rendering artifacts.</summary>
    public string? TemporaryDirectory { get; set; }
    /// <summary>Maximum Mermaid renderer execution time in milliseconds.</summary>
    public int MermaidRenderTimeoutMilliseconds { get; set; } = 30000;
    /// <summary>
    /// Optional deck preflight rules applied after conversion. When omitted, bounded preflight runs with
    /// pairwise shape-collision detection disabled; trusted callers can opt into that diagnostic explicitly.
    /// </summary>
    public PowerPointDeckPreflightOptions? PreflightOptions { get; set; }
    /// <summary>Whether preflight findings should fail the conversion.</summary>
    public bool FailOnPreflightFindings { get; set; }
}
