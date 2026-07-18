namespace OfficeIMO.Drawing;

/// <summary>Optional image-codec boundary for formats not decoded by the dependency-free Drawing core.</summary>
public interface IOfficeRasterImageCodec {
    /// <summary>Attempts to decode encoded image bytes into an RGBA raster.</summary>
    bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image);
}

/// <summary>Options for rasterizing an Office drawing with an optional image codec.</summary>
public sealed class OfficeDrawingRasterRenderOptions {
    /// <summary>Output scale. Defaults to 1.</summary>
    public double Scale { get; set; } = 1D;
    /// <summary>Optional canvas background.</summary>
    public OfficeColor? Background { get; set; }
    /// <summary>Optional decoder for formats not handled by the dependency-free core.</summary>
    public IOfficeRasterImageCodec? ImageCodec { get; set; }
}
