using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Positioned image element inside an <see cref="OfficeDrawing"/> canvas.
/// </summary>
public sealed class OfficeDrawingImage : OfficeDrawingElement {
    private readonly byte[] _bytes;

    /// <summary>
    /// Creates an image drawing element.
    /// </summary>
    /// <param name="bytes">Embedded source image bytes.</param>
    /// <param name="contentType">Optional MIME content type.</param>
    /// <param name="projection">Destination placement, crop, rotation, and flip settings.</param>
    /// <param name="alternativeText">Optional semantic label used by adapters and diagnostics.</param>
    /// <param name="opacity">Element opacity from 0 transparent to 1 opaque.</param>
    public OfficeDrawingImage(byte[] bytes, string? contentType, OfficeImageProjection projection, string? alternativeText = null, double opacity = 1D)
        : this(bytes, contentType, projection, alternativeText, opacity, useDataSnapshot: false) {
    }

    internal OfficeDrawingImage(byte[] bytes, string? contentType, OfficeImageProjection projection, string? alternativeText, double opacity, bool useDataSnapshot) {
        if (bytes == null) {
            throw new ArgumentNullException(nameof(bytes));
        }

        if (bytes.Length == 0) {
            throw new ArgumentException("Image bytes cannot be empty.", nameof(bytes));
        }

        if (double.IsNaN(opacity) || double.IsInfinity(opacity) || opacity < 0D || opacity > 1D) {
            throw new ArgumentOutOfRangeException(nameof(opacity), "Image opacity must be between 0 and 1.");
        }

        _bytes = useDataSnapshot ? bytes : (byte[])bytes.Clone();
        ContentType = OfficeImageInfo.TryNormalizeImageContentType(contentType, out string normalizedContentType)
            ? normalizedContentType
            : OfficeImageInfo.NormalizeMimeType(contentType);
        Projection = projection;
        AlternativeText = alternativeText;
        Opacity = opacity;
    }

    /// <summary>Embedded source image bytes.</summary>
    public byte[] Bytes => (byte[])_bytes.Clone();

    /// <summary>Normalized MIME content type, when available.</summary>
    public string ContentType { get; }

    /// <summary>Destination placement, crop, rotation, and flip settings.</summary>
    public OfficeImageProjection Projection { get; }

    /// <summary>Optional semantic label used by adapters and diagnostics.</summary>
    public string? AlternativeText { get; }

    /// <summary>Element opacity from 0 transparent to 1 opaque.</summary>
    public double Opacity { get; }

    internal byte[] EncodedBytes => _bytes;

    internal override OfficeDrawingElement CloneElement() =>
        new OfficeDrawingImage(_bytes, ContentType, Projection, AlternativeText, Opacity, useDataSnapshot: true);
}
