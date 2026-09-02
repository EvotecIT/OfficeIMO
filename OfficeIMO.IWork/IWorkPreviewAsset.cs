namespace OfficeIMO.IWork;

/// <summary>An embedded visual preview discovered in an iWork package.</summary>
public sealed class IWorkPreviewAsset {
    private readonly byte[] _bytes;

    internal IWorkPreviewAsset(string path, string mediaType, IWorkVisualCoverage coverage,
        int? pixelWidth, int? pixelHeight, byte[] bytes) {
        Path = path;
        MediaType = mediaType;
        Coverage = coverage;
        PixelWidth = pixelWidth;
        PixelHeight = pixelHeight;
        _bytes = bytes;
    }

    /// <summary>Gets the source package path.</summary>
    public string Path { get; }
    /// <summary>Gets the preview media type.</summary>
    public string MediaType { get; }
    /// <summary>Gets the known preview coverage.</summary>
    public IWorkVisualCoverage Coverage { get; }
    /// <summary>Gets the raster width when it can be read without decoding the image.</summary>
    public int? PixelWidth { get; }
    /// <summary>Gets the raster height when it can be read without decoding the image.</summary>
    public int? PixelHeight { get; }
    /// <summary>Gets the preview byte length.</summary>
    public int Length => _bytes.Length;
    /// <summary>Returns a defensive copy of the preview bytes.</summary>
    public byte[] GetBytes() => (byte[])_bytes.Clone();
    internal byte[] Bytes => _bytes;
}
