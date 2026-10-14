using System.Security.Cryptography;
using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.IWork.Internal;

/// <summary>Owns a validated immutable image payload shared by native document creators.</summary>
internal sealed class IWorkNativeImageData {
    internal IWorkNativeImageData(byte[] data, string? accessibilityDescription, CancellationToken cancellationToken) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        cancellationToken.ThrowIfCancellationRequested();
        if (!OfficeRasterGuards.IsEncodedPayloadWithinLimits(data.Length)
            || !OfficeImageReader.TryIdentifyByContent(data, null, cancellationToken, out OfficeImageInfo info)
            || info.Format is not OfficeImageFormat.Png and not OfficeImageFormat.Jpeg)
            throw new ArgumentException("A bounded PNG or JPEG payload is required.", nameof(data));
        EncodedBytes = (byte[])data.Clone();
        if (!OfficeImageReader.TryValidateContent(EncodedBytes, null, cancellationToken, out info)
            || info.Format is not OfficeImageFormat.Png and not OfficeImageFormat.Jpeg)
            throw new ArgumentException("A complete PNG or JPEG payload is required.", nameof(data));
        if (accessibilityDescription != null) {
            // CRLF normalization can at most halve the raw UTF-16 length.
            if (accessibilityDescription.Length > 8192)
                throw new ArgumentException("An image description supports at most 4096 UTF-16 code units.", nameof(accessibilityDescription));
            accessibilityDescription = IWorkCreationValidation.NormalizeText(accessibilityDescription);
            if (accessibilityDescription.Length > 4096)
                throw new ArgumentException("An image description supports at most 4096 UTF-16 code units.", nameof(accessibilityDescription));
        }
        PixelWidth = info.Width; PixelHeight = info.Height;
        MediaType = info.MimeType; Extension = OfficeImageInfo.GetDefaultExtension(info.Format);
        AccessibilityDescription = accessibilityDescription;
        using var hash = SHA256.Create(); ContentHash = hash.ComputeHash(EncodedBytes);
        cancellationToken.ThrowIfCancellationRequested();
    }

    internal byte[] EncodedBytes { get; }
    internal byte[] ContentHash { get; }
    internal int PixelWidth { get; }
    internal int PixelHeight { get; }
    internal string MediaType { get; }
    internal string Extension { get; }
    internal string? AccessibilityDescription { get; }
    internal long DecodedBytes => checked((long)PixelWidth * PixelHeight * 4);
}
