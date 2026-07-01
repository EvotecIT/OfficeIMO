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
    public OfficeDrawingImage(byte[] bytes, string? contentType, OfficeImageProjection projection, string? alternativeText = null) {
        if (bytes == null) {
            throw new ArgumentNullException(nameof(bytes));
        }

        if (bytes.Length == 0) {
            throw new ArgumentException("Image bytes cannot be empty.", nameof(bytes));
        }

        _bytes = (byte[])bytes.Clone();
        ContentType = OfficeImageInfo.TryNormalizeImageContentType(contentType, out string normalizedContentType)
            ? normalizedContentType
            : OfficeImageInfo.NormalizeMimeType(contentType);
        Projection = projection;
        AlternativeText = alternativeText;
    }

    /// <summary>Embedded source image bytes.</summary>
    public byte[] Bytes => (byte[])_bytes.Clone();

    /// <summary>Normalized MIME content type, when available.</summary>
    public string ContentType { get; }

    /// <summary>Destination placement, crop, rotation, and flip settings.</summary>
    public OfficeImageProjection Projection { get; }

    /// <summary>Optional semantic label used by adapters and diagnostics.</summary>
    public string? AlternativeText { get; }

    internal override OfficeDrawingElement CloneElement() =>
        new OfficeDrawingImage(_bytes, ContentType, Projection, AlternativeText);
}
