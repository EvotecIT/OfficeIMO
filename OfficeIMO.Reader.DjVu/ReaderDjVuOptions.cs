using OfficeIMO.DjVu;

namespace OfficeIMO.Reader.DjVu;

/// <summary>Which pages to decode into raster assets during DjVu ingestion.</summary>
public enum ReaderDjVuImageMode {
    /// <summary>Extract stored text and geometry without decoding page images.</summary>
    None,
    /// <summary>Decode only pages whose stored text layer is absent or empty, for optional OCR.</summary>
    MissingTextPages,
    /// <summary>Decode every page as a complete raster preview.</summary>
    AllPages
}

/// <summary>Native DjVu limits and optional page images, snapshotted when the handler is registered.</summary>
public sealed class ReaderDjVuOptions {
    /// <summary>Owned document limits and explicit indirect component resolver.</summary>
    public DjVuReadOptions ReadOptions { get; set; } = new DjVuReadOptions();
    /// <summary>Page-image policy. Text-only ingestion is the default.</summary>
    public ReaderDjVuImageMode ImageMode { get; set; }
    /// <summary>One-based source pages in requested output order. Null selects all pages; duplicates are rejected.</summary>
    public IReadOnlyList<int>? PageNumbers { get; set; }
    /// <summary>Complete-page image settings. The default is 150 DPI; regions and disabled rotation are not allowed for Reader assets.</summary>
    public DjVuRenderOptions RenderOptions { get; set; } = new DjVuRenderOptions { Dpi = 150 };
    /// <summary>Maximum emitted page images. Zero rejects any requested image; use ImageMode.None for text-only ingestion.</summary>
    public int MaxPageImages { get; set; } = 512;
    /// <summary>Maximum encoded PNG bytes per page image.</summary>
    public long MaxPageImageBytes { get; set; } = 32L * 1024 * 1024;
    /// <summary>Maximum aggregate encoded page-image bytes.</summary>
    public long MaxTotalPageImageBytes { get; set; } = 128L * 1024 * 1024;
    /// <summary>Creates an independent, validated settings snapshot.</summary>
    public ReaderDjVuOptions Clone() {
        var copy = (ReaderDjVuOptions)MemberwiseClone();
        copy.ReadOptions = (ReadOptions ?? throw new ArgumentNullException(nameof(ReadOptions))).Clone();
        copy.RenderOptions = (RenderOptions ?? throw new ArgumentNullException(nameof(RenderOptions))).Clone();
        copy.PageNumbers = PageNumbers == null ? null : Array.AsReadOnly(PageNumbers.ToArray());
        if (!Enum.IsDefined(typeof(ReaderDjVuImageMode), ImageMode)) throw new ArgumentOutOfRangeException(nameof(ImageMode));
        if (RenderOptions.Region.HasValue || !RenderOptions.ApplyRotation) throw new ArgumentException("Reader page images require the complete, display-oriented page.", nameof(RenderOptions));
        if (MaxPageImages < 0) throw new ArgumentOutOfRangeException(nameof(MaxPageImages));
        if (MaxPageImageBytes <= 0 || MaxPageImageBytes > 128L * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(MaxPageImageBytes));
        if (MaxTotalPageImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTotalPageImageBytes));
        return copy;
    }
}
