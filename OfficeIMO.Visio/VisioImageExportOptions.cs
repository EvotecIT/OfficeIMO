using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

/// <summary>Options for format-neutral dependency-free Visio image export.</summary>
public sealed class VisioImageExportOptions : OfficeImageExportOptions {
    /// <summary>Zero-based first page index used by document export.</summary>
    public int PageIndex { get; set; }

    /// <summary>Maximum number of document pages to export, or all remaining pages when absent.</summary>
    public int? PageCount { get; set; }

    /// <summary>Whether shape and connector text is rendered.</summary>
    public bool RenderText { get; set; } = true;

    /// <summary>Optional TrueType/OpenType font file used for native raster text outlines.</summary>
    public string? FontFilePath { get; set; }

    /// <summary>Optional font face name used when selecting a face from a font collection.</summary>
    public string? FontFaceName { get; set; }

    /// <summary>Optional zero-based font face index used when selecting a face from a font collection.</summary>
    public int? FontCollectionIndex { get; set; }

    /// <summary>Whether built-in OfficeIMO stencil metadata is projected as vector artwork.</summary>
    public bool RenderStencilArtwork { get; set; } = true;

    /// <summary>Whether connector labels are rendered.</summary>
    public bool RenderConnectorLabels { get; set; } = true;

    /// <summary>Whether connector labels are nudged away from page edges, shapes, and earlier labels.</summary>
    public bool ResolveConnectorLabelOverlaps { get; set; } = true;

    /// <summary>Supersampling factor used for smoother raster output.</summary>
    public int Supersampling { get; set; } = 3;

    /// <summary>Whether SVG output includes an XML declaration.</summary>
    public bool IncludeSvgXmlDeclaration { get; set; }

    /// <summary>Maximum number of pixels in a resolved raster result before scale is reduced.</summary>
    public long MaximumRasterPixels { get; set; } = 50_000_000L;

    internal VisioImageExportOptions Clone() => new VisioImageExportOptions {
        Scale = Scale,
        BackgroundColor = BackgroundColor,
        RasterEncoding = RasterEncoding?.Clone() ?? new OfficeRasterEncodingOptions(),
        PageIndex = PageIndex,
        PageCount = PageCount,
        RenderText = RenderText,
        FontFilePath = FontFilePath,
        FontFaceName = FontFaceName,
        FontCollectionIndex = FontCollectionIndex,
        RenderStencilArtwork = RenderStencilArtwork,
        RenderConnectorLabels = RenderConnectorLabels,
        ResolveConnectorLabelOverlaps = ResolveConnectorLabelOverlaps,
        Supersampling = Supersampling,
        IncludeSvgXmlDeclaration = IncludeSvgXmlDeclaration,
        MaximumRasterPixels = MaximumRasterPixels
    };

    internal void Validate() {
        OfficeImageExportOptions.ValidateScale(Scale, nameof(Scale));
        if (PageIndex < 0) throw new ArgumentOutOfRangeException(nameof(PageIndex));
        if (PageCount.HasValue && PageCount.Value < 1) throw new ArgumentOutOfRangeException(nameof(PageCount));
        if (Supersampling < 1 || Supersampling > 4) {
            throw new ArgumentOutOfRangeException(nameof(Supersampling), "Supersampling must be between 1 and 4.");
        }
        if (FontCollectionIndex.HasValue && FontCollectionIndex.Value < 0) {
            throw new ArgumentOutOfRangeException(nameof(FontCollectionIndex));
        }
        if (MaximumRasterPixels < 1L) throw new ArgumentOutOfRangeException(nameof(MaximumRasterPixels));
    }
}
