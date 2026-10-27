using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Process-wide fallback for optional system raster codecs. Explicit request codecs take precedence.</summary>
public static class OfficeRasterImageCodecs {
    private static IOfficeRasterImageCodec? _platform;

    /// <summary>Gets or replaces the optional platform decoder; null disables the fallback.</summary>
    public static IOfficeRasterImageCodec? Platform {
        get => Volatile.Read(ref _platform);
        set => Interlocked.Exchange(ref _platform, value);
    }

    internal static IOfficeRasterImageCodec? Resolve(IOfficeRasterImageCodec? codec) {
        while (true) {
            if (codec is OfficeRasterImageFallbackCodec fallback) codec = fallback.SourceCodec;
            else if (codec is OfficeDrawingRasterRenderer.RequiredImageCodec required) codec = required.SourceCodec;
            else return codec ?? Platform;
        }
    }
}

/// <summary>Why a licensed operating-system image codec cannot be used.</summary>
public enum OfficeRasterCodecUnavailableReason {
    /// <summary>The codec is available.</summary>
    None,
    /// <summary>No platform codec is registered on this host.</summary>
    NoPlatformCodec,
    /// <summary>Windows has no registered HEIF image decoder.</summary>
    WindowsHeifExtensionMissing,
    /// <summary>The registered Windows HEIF decoder cannot decode the HEVC probe.</summary>
    WindowsHevcExtensionMissing
}

/// <summary>Availability and an actionable explanation for an optional platform codec.</summary>
public sealed class OfficeRasterCodecAvailability {
    /// <summary>Creates an availability result.</summary>
    public OfficeRasterCodecAvailability(OfficeRasterCodecUnavailableReason reason, string? message = null) {
        Reason = reason;
        Message = message ?? (reason == OfficeRasterCodecUnavailableReason.None ? string.Empty :
            "HEIC requires OfficeIMO.Drawing.ImageIO on macOS/iOS or OfficeIMO.Drawing.Wic on Windows. Register the platform codec; Windows also requires Microsoft's HEIF Image Extensions and HEVC Video Extensions.");
    }
    /// <summary>Whether the licensed system codec can decode HEIC.</summary>
    public bool IsAvailable => Reason == OfficeRasterCodecUnavailableReason.None;
    /// <summary>The missing platform component.</summary>
    public OfficeRasterCodecUnavailableReason Reason { get; }
    /// <summary>An actionable explanation.</summary>
    public string Message { get; }
}

/// <summary>Optional availability reporting for a raster codec backed by installed system components.</summary>
public interface IOfficeRasterImageCodecAvailability {
    /// <summary>Checks the system decoder and its required components.</summary>
    OfficeRasterCodecAvailability GetAvailability();
}
