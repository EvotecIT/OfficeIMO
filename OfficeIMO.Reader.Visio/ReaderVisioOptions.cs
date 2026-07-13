using OfficeIMO.Visio;

namespace OfficeIMO.Reader.Visio;

/// <summary>
/// Options for Visio ingestion registered through <see cref="OfficeDocumentReaderBuilderVisioExtensions"/>.
/// </summary>
public sealed class ReaderVisioOptions {
    /// <summary>
    /// When true, emits per-page SVG preview asset metadata in read results.
    /// </summary>
    public bool IncludeSvgPreviewAssets { get; set; }

    /// <summary>
    /// When true, emits per-page PNG preview asset metadata in read results.
    /// </summary>
    public bool IncludePngPreviewAssets { get; set; }

    /// <summary>
    /// Optional SVG rendering options used when <see cref="IncludeSvgPreviewAssets"/> is true.
    /// </summary>
    public VisioSvgSaveOptions? SvgOptions { get; set; }

    /// <summary>
    /// Optional PNG rendering options used when <see cref="IncludePngPreviewAssets"/> is true.
    /// </summary>
    public VisioPngSaveOptions? PngOptions { get; set; }
}
