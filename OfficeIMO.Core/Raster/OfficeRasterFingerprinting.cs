using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Computes deterministic perceptual fingerprints from managed raster images.</summary>
public static class OfficeRasterFingerprinting {
    /// <summary>Computes the 64-bit horizontal difference hash of an image.</summary>
    /// <remarks>
    /// The image is resized to nine by eight pixels with bicubic filtering in encoded sRGB,
    /// then compared using rounded BT.709 luminance. A bit is set when the left pixel is
    /// brighter than its right neighbor. Bits run left to right, top to bottom, starting
    /// with the least significant bit. Alpha participates in premultiplied resampling;
    /// the resulting RGB luminance is compared without compositing over a background.
    /// A difference hash is a visual similarity signal, not a cryptographic digest.
    /// Recompute stored hashes when changing decoding or resampling implementations.
    /// </remarks>
    public static ulong DifferenceHash(OfficeRasterImage source, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        OfficeRasterImage thumbnail = OfficeRasterFilters.Grayscale(OfficeRasterResampler.Resize(source, 9, 8,
            OfficeRasterResamplingMode.Bicubic, OfficeRasterResamplingColorSpace.EncodedSrgb, cancellationToken),
            OfficeRasterGrayscaleMode.Bt709, cancellationToken);
        byte[] pixels = thumbnail.PixelBuffer;
        ulong hash = 0;
        for (int y = 0; y < 8; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            int row = y * 9 * 4;
            int left = pixels[row];
            for (int x = 0; x < 8; x++) {
                int right = pixels[row + (x + 1) * 4];
                if (left > right) hash |= 1UL << (y * 8 + x);
                left = right;
            }
        }
        return hash;
    }

}
