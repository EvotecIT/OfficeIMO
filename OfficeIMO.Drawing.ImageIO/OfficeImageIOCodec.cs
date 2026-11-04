using System;

namespace OfficeIMO.Drawing.ImageIO;

/// <summary>Decodes raster samples using Apple's system ImageIO, CoreGraphics and CoreFoundation frameworks.</summary>
/// <remarks>HEIC input from Core consists of one synthesized hvc1 item without container transforms or colr.
/// The returned straight-alpha RGBA samples retain the decoded source RGB encoding. Core owns composition,
/// orientation, crop and ICC/nclx conversion. This codec is synchronous and does not apply display color matching.</remarks>
public sealed class OfficeImageIOCodec : IOfficeRasterImageCodec, IOfficeRasterImageCodecAvailability {
    private const long MaximumPixels = 50_000_000;
    private const int MaximumBytes = 128 * 1024 * 1024;
    private const long WorkingSetBytes = 256L * 1024 * 1024;
    private readonly long _maximumDecodedPixels;
    private readonly int _maximumEncodedBytes;

    /// <summary>Creates a codec with bounded encoded input and decoded tile size.</summary>
    /// <param name="maximumDecodedPixels">Maximum pixels per tile, from 1 through 50,000,000.</param>
    /// <param name="maximumEncodedBytes">Maximum input bytes, from 1 through 128 MiB.</param>
    public OfficeImageIOCodec(long maximumDecodedPixels = MaximumPixels, int maximumEncodedBytes = MaximumBytes) {
        if (maximumDecodedPixels < 1 || maximumDecodedPixels > MaximumPixels)
            throw new ArgumentOutOfRangeException(nameof(maximumDecodedPixels));
        if (maximumEncodedBytes < 1 || maximumEncodedBytes > MaximumBytes)
            throw new ArgumentOutOfRangeException(nameof(maximumEncodedBytes));
        _maximumDecodedPixels = maximumDecodedPixels;
        _maximumEncodedBytes = maximumEncodedBytes;
    }

    /// <summary>Whether this process runs on macOS 10.13+, iOS 11+, or Mac Catalyst 13.1+.</summary>
    /// <remarks>This describes the platform, not support for every HEVC profile on that device.</remarks>
    public static bool IsSupported => OperatingSystem.IsMacOSVersionAtLeast(10, 13) ||
        OperatingSystem.IsIOSVersionAtLeast(11) || OperatingSystem.IsMacCatalystVersionAtLeast(13, 1);

    /// <summary>Installs this system codec as the shared platform fallback on supported Apple systems.</summary>
    /// <remarks>Does nothing on other operating systems. Explicit per-call codecs take precedence in Core.</remarks>
    public static void Register() {
        if (IsSupported) OfficeRasterImageCodecs.Platform = new OfficeImageIOCodec();
    }

    /// <summary>Reports whether the Apple system codec is available on this operating system.</summary>
    /// <remarks>Individual bit depths and HEVC profiles still depend on the installed OS and device decoder.</remarks>
    public OfficeRasterCodecAvailability GetAvailability() => new OfficeRasterCodecAvailability(
        IsSupported ? OfficeRasterCodecUnavailableReason.None : OfficeRasterCodecUnavailableReason.NoPlatformCodec,
        IsSupported ? null : "Apple ImageIO HEIC decoding requires macOS 10.13+, iOS 11+, or Mac Catalyst 13.1+.");

    /// <summary>Decodes the first stored image to top-to-bottom, 8-bit straight-alpha RGBA in source RGB space.</summary>
    /// <remarks>The content type is advisory; ImageIO identifies the input bytes. Invalid, unsupported or
    /// over-budget input returns false and a null image. Full HEIF files must go through Core's composer.</remarks>
    public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
        image = null;
        if (!IsSupported || encodedBytes == null || encodedBytes.Length == 0 ||
            encodedBytes.Length > _maximumEncodedBytes || encodedBytes.LongLength * 2 + 65_536 > WorkingSetBytes)
            return false;
        try {
            return ImageIODecoder.TryDecode(encodedBytes, _maximumDecodedPixels, WorkingSetBytes, out image);
        } catch (DllNotFoundException) {
            return false;
        } catch (EntryPointNotFoundException) {
            return false;
        }
    }
}
