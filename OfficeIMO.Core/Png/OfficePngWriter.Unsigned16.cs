using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficePngWriter {
    /// <summary>Encodes tightly packed straight-alpha 16-bit RGBA samples without reducing precision.</summary>
    /// <remarks>Samples are in host ushort order. PNG stores them in network byte order.
    /// Exact gray and opaque channels are reduced to gray, gray-alpha, or RGB while
    /// retaining 16-bit depth. This entry point uses optimal adaptive compression.
    /// The 8-bit OfficeRasterImage model cannot retain 16-bit source precision.</remarks>
    public static byte[] EncodeRgba64(int width, int height, ushort[] rgba,
        OfficePngEncodeOptions? options = null, CancellationToken cancellationToken = default) {
        if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
        if (rgba == null) throw new ArgumentNullException(nameof(rgba));
        if (rgba.LongLength != checked((long)width * height * 4)) throw new ArgumentException("RGBA sample count must match the dimensions.", nameof(rgba));
        options ??= new OfficePngEncodeOptions();
        ValidatePngEffort(options.Effort);
        if (options.Compression != OfficePngCompression.Optimal) throw new ArgumentException("16-bit encoding requires optimal compression.", nameof(options));
        ValidateDpi(options.DpiX, nameof(options.DpiX)); ValidateDpi(options.DpiY, nameof(options.DpiY));
        OfficeRasterGuards.EnsureOutputPixels(width, height, "PNG dimensions exceed the managed image limit.");
        if (rgba.LongLength * 2 + (long)width * 32 + 65536 > OfficeRasterGuards.MaximumDecodedBytes)
            throw new ArgumentException("16-bit PNG encoding exceeds the managed working-set limit.", nameof(rgba));
        bool gray = true, opaque = true;
        for (int i = 0; i < rgba.Length; i += 4) {
            if ((i & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            gray &= rgba[i] == rgba[i + 1] && rgba[i] == rgba[i + 2]; opaque &= rgba[i + 3] == ushort.MaxValue;
        }
        var layout = new PngColorLayout(gray ? opaque ? 0 : 4 : opaque ? 2 : 6, 16);
        using var output = new MemoryStream();
        WriteColorPng(output, width, height, null, rgba, layout, true,
            options.WritePhysicalResolution ? options.DpiX : (double?)null,
            options.WritePhysicalResolution ? options.DpiY : (double?)null,
            cancellationToken, null,
            level: options.Effort == OfficeImageEncodingEffort.Smallest ? null : options.Effort == OfficeImageEncodingEffort.Fast ? System.IO.Compression.CompressionLevel.Fastest : System.IO.Compression.CompressionLevel.Optimal,
            sampledFilters: options.Effort != OfficeImageEncodingEffort.Smallest);
        return output.ToArray();
    }
}
