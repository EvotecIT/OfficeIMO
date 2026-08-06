namespace OfficeIMO.Drawing;

/// <summary>PNG encoding options.</summary>
public sealed class OfficePngEncodeOptions {
    /// <summary>Zlib compression strategy.</summary>
    public OfficePngCompression Compression { get; set; } = OfficePngCompression.Optimal;

    /// <summary>Horizontal resolution written to the PNG pHYs chunk.</summary>
    public double DpiX { get; set; } = 96D;

    /// <summary>Vertical resolution written to the PNG pHYs chunk.</summary>
    public double DpiY { get; set; } = 96D;
}
