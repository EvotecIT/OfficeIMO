using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Built-in managed routes for an image container, subject to the documented subset and resource limits.</summary>
/// <remarks>Capabilities describe available routes, not a guarantee that every variant or payload in a format decodes.
/// Caller-supplied codecs do not change this built-in descriptor.</remarks>
public sealed class OfficeRasterFormatCapabilities {
    internal OfficeRasterFormatCapabilities(OfficeImageFormat format, bool inspect, bool decode, bool frames, bool encode,
        bool encodeFrames, bool readMetadata, OfficeImageMetadataProfileKinds profiles, string subset) {
        Format = format; CanInspect = inspect; CanDecode = decode; CanDecodeMultipleFrames = frames;
        CanEncode = encode; CanEncodeMultipleFrames = encodeFrames; CanReadMetadata = readMetadata;
        SupportedMetadataProfiles = profiles; DecodeSubset = subset;
    }
    /// <summary>The encoded container family.</summary>
    public OfficeImageFormat Format { get; }
    /// <summary>Whether the content reader has a format-identification route.</summary>
    public bool CanIdentify => Format != OfficeImageFormat.Unknown;
    /// <summary>Whether the managed raster inspector supports source frame or page descriptors.</summary>
    public bool CanInspect { get; }
    /// <summary>Whether built-in pixel decoding supports the subset described by DecodeSubset.</summary>
    public bool CanDecode { get; }
    /// <summary>Whether built-in decoding supports multiple source frames, pages, or icon entries.</summary>
    public bool CanDecodeMultipleFrames { get; }
    /// <summary>Whether the shared raster encoder supports still output in this container family.</summary>
    public bool CanEncode { get; }
    /// <summary>Whether the shared still encoder supports multiple TIFF pages or icon entries.</summary>
    /// <remarks>Animated encoding is provided by the animation engine.</remarks>
    public bool CanEncodeMultipleFrames { get; }
    /// <summary>Whether the portable metadata reader extracts supported primary-image metadata.</summary>
    public bool CanReadMetadata { get; }
    /// <summary>Whether portable metadata can be applied without recompressing pixels.</summary>
    public bool CanApplyMetadata => SupportedMetadataProfiles != OfficeImageMetadataProfileKinds.None;
    /// <summary>Profile families supported by the portable metadata application route.</summary>
    public OfficeImageMetadataProfileKinds SupportedMetadataProfiles { get; }
    /// <summary>Scope of the built-in pixel route; identification alone does not imply pixel decoding.</summary>
    public string DecodeSubset { get; }
}

/// <summary>Truthful discovery of existing managed raster routes over the shared image format enum.</summary>
public static class OfficeRasterImageFormats {
    private static readonly IReadOnlyList<OfficeRasterFormatCapabilities> Formats = CreateFormats();

    /// <summary>Descriptors for every shared image format, including identification-only formats and Unknown.</summary>
    public static IReadOnlyList<OfficeRasterFormatCapabilities> BuiltInFormats => Formats;

    /// <summary>Returns the built-in routes for a defined image format.</summary>
    public static OfficeRasterFormatCapabilities GetCapabilities(OfficeImageFormat format) {
        foreach (OfficeRasterFormatCapabilities descriptor in Formats) if (descriptor.Format == format) return descriptor;
        throw new ArgumentOutOfRangeException(nameof(format));
    }

    private static IReadOnlyList<OfficeRasterFormatCapabilities> CreateFormats() {
        var result = new List<OfficeRasterFormatCapabilities>();
        foreach (OfficeImageFormat format in Enum.GetValues(typeof(OfficeImageFormat))) {
            bool decode = format == OfficeImageFormat.Png || format == OfficeImageFormat.Jpeg || format == OfficeImageFormat.Gif ||
                format == OfficeImageFormat.Bmp || format == OfficeImageFormat.Tiff || format == OfficeImageFormat.Webp ||
                format == OfficeImageFormat.JpegXr || format == OfficeImageFormat.Avif || format == OfficeImageFormat.Icon ||
                format == OfficeImageFormat.PortableMap || format == OfficeImageFormat.Tga;
            bool frames = format == OfficeImageFormat.Png || format == OfficeImageFormat.Gif || format == OfficeImageFormat.Tiff || format == OfficeImageFormat.Icon;
            bool encode = format == OfficeImageFormat.Png || format == OfficeImageFormat.Jpeg || format == OfficeImageFormat.Tiff ||
                format == OfficeImageFormat.Webp || format == OfficeImageFormat.Bmp || format == OfficeImageFormat.PortableMap || format == OfficeImageFormat.Tga || format == OfficeImageFormat.Icon;
            OfficeImageMetadataProfileKinds profiles = format switch {
                OfficeImageFormat.Jpeg or OfficeImageFormat.Tiff => OfficeImageMetadataProfileKinds.Exif | OfficeImageMetadataProfileKinds.Xmp | OfficeImageMetadataProfileKinds.Icc | OfficeImageMetadataProfileKinds.Iptc,
                OfficeImageFormat.Png or OfficeImageFormat.Webp => OfficeImageMetadataProfileKinds.Exif | OfficeImageMetadataProfileKinds.Xmp | OfficeImageMetadataProfileKinds.Icc,
                OfficeImageFormat.Gif => OfficeImageMetadataProfileKinds.Xmp | OfficeImageMetadataProfileKinds.Icc,
                OfficeImageFormat.Bmp => OfficeImageMetadataProfileKinds.Icc,
                _ => OfficeImageMetadataProfileKinds.None
            };
            string subset = format switch {
                OfficeImageFormat.Png => "Validated PNG samples and composed APNG frames.",
                OfficeImageFormat.Jpeg => "Managed JPEG sample and scan subset; orientation normalization is optional.",
                OfficeImageFormat.Gif => "Validated GIF frames composed on the logical canvas with stored timing.",
                OfficeImageFormat.Bmp => "Managed uncompressed Windows bitmap subset.",
                OfficeImageFormat.Tiff => "Bounded classic TIFF pages, supported sample/compression layouts, and optional orientation normalization.",
                OfficeImageFormat.Webp => "Static VP8L and VP8 with supported ALPH. Animated pixels require a caller codec.",
                OfficeImageFormat.JpegXr => "Unsigned eight/sixteen-bit and finite fixed/floating-point gray/RGB subset.",
                OfficeImageFormat.Avif => "Bounded eight/ten-bit YUV420 or monochrome still items with optional straight alpha.",
                OfficeImageFormat.Icon => "Bounded PNG-compressed and supported DIB icon entries.",
                OfficeImageFormat.PortableMap => "Portable bitmap, graymap, and pixmap subset; output is binary bilevel PBM.",
                OfficeImageFormat.Tga => "Managed TGA raster subset; output is uncompressed 32-bit straight alpha.",
                _ => "No built-in managed raster pixel route."
            };
            result.Add(new OfficeRasterFormatCapabilities(format, decode, decode, frames, encode,
                format == OfficeImageFormat.Tiff || format == OfficeImageFormat.Icon,
                profiles != OfficeImageMetadataProfileKinds.None || format == OfficeImageFormat.JpegXr, profiles, subset));
        }
        return result.AsReadOnly();
    }
}
