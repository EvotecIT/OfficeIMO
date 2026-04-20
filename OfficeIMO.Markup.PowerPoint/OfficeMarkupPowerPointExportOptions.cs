namespace OfficeIMO.Markup.PowerPoint;

public sealed class OfficeMarkupPowerPointExportOptions {
    public string OutputPath { get; set; } = string.Empty;
    public string? BaseDirectory { get; set; }
    public double SlideWidthInches { get; set; } = 10.0;
    public double SlideHeightInches { get; set; } = 5.625;
    public bool IncludeUnsupportedBlocksAsText { get; set; } = true;
    public bool RenderMermaidDiagrams { get; set; } = true;
    public string? MermaidRendererPath { get; set; }
    public string? TemporaryDirectory { get; set; }
    public int MermaidRenderTimeoutMilliseconds { get; set; } = 30000;
}
