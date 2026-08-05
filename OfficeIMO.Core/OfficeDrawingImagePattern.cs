using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Clipped repeating image pattern inside an <see cref="OfficeDrawing" /> canvas.
/// </summary>
public sealed class OfficeDrawingImagePattern : OfficeDrawingElement {
    private readonly byte[] _bytes;

    /// <summary>Creates a bounded image-pattern element.</summary>
    public OfficeDrawingImagePattern(
        byte[] bytes,
        string? contentType,
        OfficeImagePatternLayout layout,
        int maximumTileCount = 16384,
        double opacity = 1D) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        if (bytes.Length == 0) throw new ArgumentException("Image-pattern bytes cannot be empty.", nameof(bytes));
        if (maximumTileCount <= 0) throw new ArgumentOutOfRangeException(nameof(maximumTileCount), "Maximum image-pattern tile count must be positive.");
        if (layout.EstimatedTileCount > maximumTileCount) throw new ArgumentException("Image pattern exceeds the configured tile-count limit.", nameof(layout));
        if (double.IsNaN(opacity) || double.IsInfinity(opacity) || opacity < 0D || opacity > 1D) {
            throw new ArgumentOutOfRangeException(nameof(opacity), "Image-pattern opacity must be between 0 and 1.");
        }

        _bytes = (byte[])bytes.Clone();
        ContentType = OfficeImageInfo.TryNormalizeImageContentType(contentType, out string normalizedContentType)
            ? normalizedContentType
            : OfficeImageInfo.NormalizeMimeType(contentType);
        Layout = layout;
        MaximumTileCount = maximumTileCount;
        Opacity = opacity;
    }

    internal OfficeDrawingImagePattern(
        byte[] bytes,
        string contentType,
        OfficeImagePatternLayout layout,
        int maximumTileCount,
        double opacity,
        bool useSnapshot) {
        _bytes = useSnapshot ? bytes : (byte[])bytes.Clone();
        ContentType = contentType;
        Layout = layout;
        MaximumTileCount = maximumTileCount;
        Opacity = opacity;
    }

    /// <summary>Detached encoded source-image bytes.</summary>
    public byte[] Bytes => (byte[])_bytes.Clone();

    /// <summary>Normalized image media type, when available.</summary>
    public string ContentType { get; }

    /// <summary>Pattern area, origin tile, and repeat axes.</summary>
    public OfficeImagePatternLayout Layout { get; }

    /// <summary>Maximum tile count accepted by expansion-based backends.</summary>
    public int MaximumTileCount { get; }

    /// <summary>Element opacity from 0 transparent to 1 opaque.</summary>
    public double Opacity { get; }

    internal byte[] EncodedBytes => _bytes;

    internal override OfficeDrawingElement CloneElement() =>
        new OfficeDrawingImagePattern(_bytes, ContentType, Layout, MaximumTileCount, Opacity, useSnapshot: true);
}
