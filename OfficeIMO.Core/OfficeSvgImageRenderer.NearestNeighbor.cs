using System;
using System.Threading;

namespace OfficeIMO.Drawing;

// Retained for the existing HTML/PDF adapter's fallback catch. Nearest-neighbour
// SVG export embeds images and no longer raises a vectorization limit.
internal sealed class OfficeSvgImageVectorizationLimitException : InvalidOperationException {
    internal OfficeSvgImageVectorizationLimitException()
        : base("SVG nearest-neighbor image export exceeds the supported vectorization limit.") { }
}


public static partial class OfficeSvgImageRenderer {
    // Keep the nearest-neighbor decoder's admission, static-frame selection,
    // caller-codec provenance and placeholder contracts. Only the final SVG
    // representation changes from pixel rectangles to one embedded image.
    internal static bool TryCreateNearestDataUri(string? contentType, byte[] bytes,
        IOfficeRasterImageCodec? imageCodec, int maximumCharacters, CancellationToken token, out string uri) {
        uri = string.Empty;
        if (!TryDecodeRasterForSvg(bytes, contentType, imageCodec, token, out OfficeRasterImage? raster,
                out bool useDecodedPayload) || raster == null) return false;
        if (!useDecodedPayload && TryResolveEmbeddableContentType(contentType, bytes, null, out string nativeType)) {
            uri = CreateBoundedDataUri(nativeType, bytes, maximumCharacters, token);
        } else {
            byte[] png = OfficePngWriter.Encode(raster, token);
            uri = CreateBoundedDataUri("image/png", png, maximumCharacters, token);
        }
        return true;
    }
}
