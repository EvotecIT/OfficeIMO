namespace OfficeIMO.Visio;

/// <summary>Controls projection of a Visio document into the neutral OfficeIMO document model.</summary>
public sealed class VisioDocumentProjectionOptions {
    /// <summary>Maximum shape-data rows retained per page. The default retains every row.</summary>
    public int MaxTableRows { get; set; } = int.MaxValue;

    /// <summary>When true, includes an SVG preview asset for every projected page.</summary>
    public bool IncludeSvgPreviewAssets { get; set; }

    /// <summary>When true, includes a PNG preview asset for every projected page.</summary>
    public bool IncludePngPreviewAssets { get; set; }

    /// <summary>SVG rendering options used when SVG previews are enabled.</summary>
    public VisioSvgSaveOptions? SvgOptions { get; set; }

    /// <summary>PNG rendering options used when PNG previews are enabled.</summary>
    public VisioPngSaveOptions? PngOptions { get; set; }

    internal VisioDocumentProjectionOptions Snapshot() {
        if (MaxTableRows <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTableRows));
        return new VisioDocumentProjectionOptions {
            MaxTableRows = MaxTableRows,
            IncludeSvgPreviewAssets = IncludeSvgPreviewAssets,
            IncludePngPreviewAssets = IncludePngPreviewAssets,
            SvgOptions = SvgOptions,
            PngOptions = PngOptions
        };
    }
}
