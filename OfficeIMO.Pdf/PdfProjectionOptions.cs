using OfficeIMO;

namespace OfficeIMO.Pdf;

/// <summary>How normalized source pages are mapped into PDF pagination.</summary>
public enum PdfProjectionPagePolicy {
    /// <summary>Start a new PDF page between normalized source pages.</summary>
    PreserveSourcePages,
    /// <summary>Compose all normalized pages into one continuous PDF flow.</summary>
    ContinuousFlow
}

/// <summary>How normalized assets are represented in PDF output.</summary>
public enum PdfProjectionAssetPolicy {
    /// <summary>Embed supported raster images and list non-image resources.</summary>
    EmbedSupportedImages,
    /// <summary>List asset metadata without embedding payloads.</summary>
    ListMetadata,
    /// <summary>Omit assets with explicit conversion diagnostics.</summary>
    Omit
}

/// <summary>How normalized links are represented in PDF output.</summary>
public enum PdfProjectionLinkPolicy {
    /// <summary>Emit URI links and list navigation targets that cannot be preserved directly.</summary>
    PreserveUriLinks,
    /// <summary>List link metadata as text.</summary>
    ListMetadata,
    /// <summary>Omit links with explicit conversion diagnostics.</summary>
    Omit
}

/// <summary>How normalized source forms are represented in PDF output.</summary>
public enum PdfProjectionFormPolicy {
    /// <summary>Render field names and current values as non-interactive content.</summary>
    RenderCurrentValues,
    /// <summary>Omit source forms with explicit conversion diagnostics.</summary>
    Omit
}

/// <summary>
/// Explicit, source-neutral policy for projecting an <see cref="OfficeDocumentModel"/> into PDF.
/// Email attachment, EPUB resource/pagination, and diagram-page decisions all flow through these options.
/// </summary>
public sealed class PdfProjectionOptions {
    /// <summary>PDF generation options. The converter snapshots this value.</summary>
    public OfficeIMO.Pdf.PdfOptions? PdfOptions { get; set; }

    /// <summary>Normalized page handling.</summary>
    public PdfProjectionPagePolicy PagePolicy { get; set; } = PdfProjectionPagePolicy.PreserveSourcePages;

    /// <summary>Asset and attachment handling.</summary>
    public PdfProjectionAssetPolicy AssetPolicy { get; set; } = PdfProjectionAssetPolicy.EmbedSupportedImages;

    /// <summary>
    /// Shared raster decode settings used when assets require normalization.
    /// The converter snapshots all settings and observes their cancellation token together with the projection token.
    /// Normalization uses the smaller of the caller pixel limit and the shared PDF transcode ceiling.
    /// A null value uses the shared decoder defaults.
    /// </summary>
    public OfficeIMO.Drawing.OfficeRasterDecodeOptions? RasterDecodeOptions { get; set; } = new OfficeIMO.Drawing.OfficeRasterDecodeOptions();

    /// <summary>URI and navigation handling.</summary>
    public PdfProjectionLinkPolicy LinkPolicy { get; set; } = PdfProjectionLinkPolicy.PreserveUriLinks;

    /// <summary>Source form handling.</summary>
    public PdfProjectionFormPolicy FormPolicy { get; set; } = PdfProjectionFormPolicy.RenderCurrentValues;

    /// <summary>When true, source metadata is emitted as a compact facts table.</summary>
    public bool IncludeMetadata { get; set; } = true;

    internal void Validate() {
        if (PagePolicy < PdfProjectionPagePolicy.PreserveSourcePages || PagePolicy > PdfProjectionPagePolicy.ContinuousFlow) throw new ArgumentOutOfRangeException(nameof(PagePolicy));
        if (AssetPolicy < PdfProjectionAssetPolicy.EmbedSupportedImages || AssetPolicy > PdfProjectionAssetPolicy.Omit) throw new ArgumentOutOfRangeException(nameof(AssetPolicy));
        if (LinkPolicy < PdfProjectionLinkPolicy.PreserveUriLinks || LinkPolicy > PdfProjectionLinkPolicy.Omit) throw new ArgumentOutOfRangeException(nameof(LinkPolicy));
        if (FormPolicy < PdfProjectionFormPolicy.RenderCurrentValues || FormPolicy > PdfProjectionFormPolicy.Omit) throw new ArgumentOutOfRangeException(nameof(FormPolicy));
        if (RasterDecodeOptions != null &&
            RasterDecodeOptions.FrameLossPolicy != OfficeIMO.Drawing.OfficeRasterFrameLossPolicy.UseSelectedFrame &&
            RasterDecodeOptions.FrameLossPolicy != OfficeIMO.Drawing.OfficeRasterFrameLossPolicy.RejectMultipleFrames) {
            throw new ArgumentOutOfRangeException(nameof(RasterDecodeOptions));
        }
    }

    internal OfficeIMO.Drawing.OfficeRasterDecodeOptions SnapshotRasterDecodeOptions() =>
        RasterDecodeOptions?.Clone() ?? new OfficeIMO.Drawing.OfficeRasterDecodeOptions();
}
