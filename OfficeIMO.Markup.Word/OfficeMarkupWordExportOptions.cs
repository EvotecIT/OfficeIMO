namespace OfficeIMO.Markup.Word;

public sealed class OfficeMarkupWordExportOptions {
    public string OutputPath { get; set; } = string.Empty;
    public string? BaseDirectory { get; set; }
    public bool IncludeUnsupportedBlocksAsText { get; set; } = true;
    public int DefaultChartWidthPixels { get; set; } = 640;
    public int DefaultChartHeightPixels { get; set; } = 360;
}
