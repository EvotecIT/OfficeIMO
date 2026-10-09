using System.Security.Cryptography;
using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.IWork;

/// <summary>An immutable positioned PNG or JPEG in a newly created Keynote presentation.</summary>
public sealed class IWorkKeynoteImage : IWorkKeynoteCreationDrawable {
    private readonly byte[] _bytes;

    internal IWorkKeynoteImage(byte[] data, float left, float top, float width, float height,
        string? accessibilityDescription, CancellationToken cancellationToken) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        IWorkKeynoteTextBox.ValidateFinite(left, nameof(left));
        IWorkKeynoteTextBox.ValidateFinite(top, nameof(top));
        IWorkKeynoteTextBox.ValidatePositive(width, nameof(width));
        IWorkKeynoteTextBox.ValidatePositive(height, nameof(height));
        cancellationToken.ThrowIfCancellationRequested();
        // Core rejects incomplete image bodies and imposes its encoded/pixel/decode bounds.
        // Check metadata before copying an unbounded caller buffer.
        if (!OfficeRasterGuards.IsEncodedPayloadWithinLimits(data.Length)
            || !OfficeImageReader.TryIdentifyByContent(data, null, cancellationToken, out OfficeImageInfo info)
            || info.Format is not OfficeImageFormat.Png and not OfficeImageFormat.Jpeg)
            throw new ArgumentException("A bounded PNG or JPEG payload is required.", nameof(data));
        _bytes = (byte[])data.Clone();
        if (!OfficeImageReader.TryValidateContent(_bytes, null, cancellationToken, out info)
            || info.Format is not OfficeImageFormat.Png and not OfficeImageFormat.Jpeg)
            throw new ArgumentException("A complete PNG or JPEG payload is required.", nameof(data));
        if (accessibilityDescription != null) {
            // CRLF normalization can at most halve the raw UTF-16 length.
            if (accessibilityDescription.Length > 8192)
                throw new ArgumentException("An image description supports at most 4096 UTF-16 code units.", nameof(accessibilityDescription));
            accessibilityDescription = IWorkKeynoteTextBox.NormalizeText(accessibilityDescription);
            if (accessibilityDescription.Length > 4096)
                throw new ArgumentException("An image description supports at most 4096 UTF-16 code units.", nameof(accessibilityDescription));
        }
        Geometry = new IWorkGeometry(left, top, width, height, 0);
        PixelWidth = info.Width; PixelHeight = info.Height;
        MediaType = info.MimeType;
        Extension = OfficeImageInfo.GetDefaultExtension(info.Format);
        AccessibilityDescription = accessibilityDescription;
        using var hash = SHA256.Create();
        ContentHash = hash.ComputeHash(_bytes);
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>Gets the frame geometry in points. The full image fills this frame without a crop or mask.</summary>
    public IWorkGeometry Geometry { get; }
    /// <summary>Gets the validated pixel width.</summary>
    public int PixelWidth { get; }
    /// <summary>Gets the validated pixel height.</summary>
    public int PixelHeight { get; }
    /// <summary>Gets image/png or image/jpeg.</summary>
    public string MediaType { get; }
    /// <summary>Gets the encoded payload size in bytes.</summary>
    public int Length => _bytes.Length;
    /// <summary>Gets the optional native accessibility description.</summary>
    public string? AccessibilityDescription { get; }
    /// <summary>Returns a copy of the validated encoded payload.</summary>
    public byte[] GetBytes() => (byte[])_bytes.Clone();
    internal byte[] EncodedBytes => _bytes;
    internal byte[] ContentHash { get; }
    internal string Extension { get; }
    internal long DecodedBytes => checked((long)PixelWidth * PixelHeight * 4);
}
