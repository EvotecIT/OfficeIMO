using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgImageRenderer {
    /// <summary>Uses inspected source pixels before considering a final SVG placeholder or custom-format codec.</summary>
    internal static bool TryDecodeRasterForSvg(byte[] bytes, string? contentType,
        IOfficeRasterImageCodec? imageCodec, CancellationToken cancellationToken, out OfficeRasterImage? raster) {
        var fallback = imageCodec as OfficeRasterImageFallbackCodec;
        var options = new OfficeRasterDecodeOptions {
            ImageCodec = fallback != null ? fallback.SourceCodec : imageCodec,
            CancellationToken = cancellationToken
        };
        if (OfficeRasterImageDecoder.TryDecode(bytes, options, out raster, out OfficeRasterDecodeInfo info) && raster != null) {
            if (info.AnimationDiscarded || info.FramesOrPagesDiscarded) fallback?.AddStaticFrameDiagnostic(info);
            if (info.UsedCallerCodec) fallback?.AddCallerCodecDiagnostic(contentType);
            return true;
        }
        cancellationToken.ThrowIfCancellationRequested();
        // Placeholders represent loss, not decoded source pixels. Their dimensions
        // must never participate in the shared container's source validation.
        if (fallback != null && info.Container != null) {
            raster = fallback.CreateFallbackImage(contentType, Math.Min(32, info.Container.CanvasWidth),
                Math.Min(32, info.Container.CanvasHeight), info.Diagnostic);
            return true;
        }
        if (!OfficeRasterImageDecoder.CanUseUninspectedCallerCodec(bytes, options, info) || imageCodec == null)
            return false;
        bool decoded = imageCodec.TryDecode((byte[])bytes.Clone(), contentType, out raster) && raster != null;
        cancellationToken.ThrowIfCancellationRequested();
        return decoded;
    }
}
