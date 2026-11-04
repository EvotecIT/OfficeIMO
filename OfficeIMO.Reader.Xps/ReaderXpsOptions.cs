using OfficeIMO.Xps;

namespace OfficeIMO.Reader.Xps;

/// <summary>Native package limits and optional strict SVG previews for XPS ingestion.</summary>
public sealed class ReaderXpsOptions {
    /// <summary>Native package and XML limits. Snapshotted when the handler is registered.</summary>
    public XpsLoadOptions ReadOptions { get; set; } = new XpsLoadOptions();
    /// <summary>Includes self-contained page SVG assets. Unsupported rendering fails instead of returning an incomplete preview.</summary>
    public bool IncludeSvgPreviewAssets { get; set; }
    internal ReaderXpsOptions Clone() => new() { ReadOptions = (ReadOptions ?? throw new ArgumentNullException(nameof(ReadOptions))).Clone(), IncludeSvgPreviewAssets = IncludeSvgPreviewAssets };
}
