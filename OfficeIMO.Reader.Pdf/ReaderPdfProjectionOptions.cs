namespace OfficeIMO.Reader;

/// <summary>How normalized source pages are mapped into PDF pagination.</summary>
public enum ReaderPdfPagePolicy {
    /// <summary>Start a new PDF page between normalized source pages.</summary>
    PreserveSourcePages,
    /// <summary>Compose all normalized pages into one continuous PDF flow.</summary>
    ContinuousFlow
}

/// <summary>How normalized assets are represented in PDF output.</summary>
public enum ReaderPdfAssetPolicy {
    /// <summary>Embed supported raster images and list non-image resources.</summary>
    EmbedSupportedImages,
    /// <summary>List asset metadata without embedding payloads.</summary>
    ListMetadata,
    /// <summary>Omit assets with explicit conversion diagnostics.</summary>
    Omit
}

/// <summary>How normalized links are represented in PDF output.</summary>
public enum ReaderPdfLinkPolicy {
    /// <summary>Emit URI links and list navigation targets that cannot be preserved directly.</summary>
    PreserveUriLinks,
    /// <summary>List link metadata as text.</summary>
    ListMetadata,
    /// <summary>Omit links with explicit conversion diagnostics.</summary>
    Omit
}

/// <summary>How normalized source forms are represented in PDF output.</summary>
public enum ReaderPdfFormPolicy {
    /// <summary>Render field names and current values as non-interactive content.</summary>
    RenderCurrentValues,
    /// <summary>Omit source forms with explicit conversion diagnostics.</summary>
    Omit
}

/// <summary>
/// Explicit, source-neutral policy for projecting an <see cref="OfficeDocumentReadResult"/> into PDF.
/// Email attachment, EPUB resource/pagination, and diagram-page decisions all flow through these options.
/// </summary>
public sealed class ReaderPdfProjectionOptions {
    /// <summary>PDF generation options. The converter snapshots this value.</summary>
    public OfficeIMO.Pdf.PdfOptions? PdfOptions { get; set; }

    /// <summary>Normalized page handling.</summary>
    public ReaderPdfPagePolicy PagePolicy { get; set; } = ReaderPdfPagePolicy.PreserveSourcePages;

    /// <summary>Asset and attachment handling.</summary>
    public ReaderPdfAssetPolicy AssetPolicy { get; set; } = ReaderPdfAssetPolicy.EmbedSupportedImages;

    /// <summary>
    /// Shared Drawing frame-selection and animation-loss policy used when raster assets require normalization.
    /// The converter snapshots this value. A null value uses the Drawing defaults.
    /// </summary>
    public OfficeIMO.Drawing.OfficeRasterDecodeOptions? RasterDecodeOptions { get; set; } = new OfficeIMO.Drawing.OfficeRasterDecodeOptions();

    /// <summary>URI and navigation handling.</summary>
    public ReaderPdfLinkPolicy LinkPolicy { get; set; } = ReaderPdfLinkPolicy.PreserveUriLinks;

    /// <summary>Source form handling.</summary>
    public ReaderPdfFormPolicy FormPolicy { get; set; } = ReaderPdfFormPolicy.RenderCurrentValues;

    /// <summary>When true, source metadata is emitted as a compact facts table.</summary>
    public bool IncludeMetadata { get; set; } = true;

    internal void Validate() {
        if (PagePolicy < ReaderPdfPagePolicy.PreserveSourcePages || PagePolicy > ReaderPdfPagePolicy.ContinuousFlow) throw new ArgumentOutOfRangeException(nameof(PagePolicy));
        if (AssetPolicy < ReaderPdfAssetPolicy.EmbedSupportedImages || AssetPolicy > ReaderPdfAssetPolicy.Omit) throw new ArgumentOutOfRangeException(nameof(AssetPolicy));
        if (LinkPolicy < ReaderPdfLinkPolicy.PreserveUriLinks || LinkPolicy > ReaderPdfLinkPolicy.Omit) throw new ArgumentOutOfRangeException(nameof(LinkPolicy));
        if (FormPolicy < ReaderPdfFormPolicy.RenderCurrentValues || FormPolicy > ReaderPdfFormPolicy.Omit) throw new ArgumentOutOfRangeException(nameof(FormPolicy));
        if (RasterDecodeOptions != null &&
            RasterDecodeOptions.AnimationPolicy != OfficeIMO.Drawing.OfficeRasterAnimationPolicy.UseSelectedFrame &&
            RasterDecodeOptions.AnimationPolicy != OfficeIMO.Drawing.OfficeRasterAnimationPolicy.RejectAnimated) {
            throw new ArgumentOutOfRangeException(nameof(RasterDecodeOptions));
        }
    }

    internal OfficeIMO.Drawing.OfficeRasterDecodeOptions SnapshotRasterDecodeOptions() =>
        new OfficeIMO.Drawing.OfficeRasterDecodeOptions {
            FrameIndex = RasterDecodeOptions?.FrameIndex ?? 0,
            AnimationPolicy = RasterDecodeOptions?.AnimationPolicy ?? OfficeIMO.Drawing.OfficeRasterAnimationPolicy.UseSelectedFrame
        };
}
