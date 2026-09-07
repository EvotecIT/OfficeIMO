using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    // The fallback travels with the codec through nested scenes. Built-in decoding still happens
    // once in TryDecodeImage; only unsupported inputs reach this required-decoding boundary.
    private sealed class RequiredImageCodec : IOfficeRasterImageCodec {
        private readonly IOfficeRasterImageCodec? _codec;
        private readonly long _maximumPixels;
        private readonly CancellationToken _token;

        internal RequiredImageCodec(IOfficeRasterImageCodec? codec, long maximumPixels, CancellationToken token) {
            _codec = codec;
            _maximumPixels = maximumPixels;
            _token = token;
        }

        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            _token.ThrowIfCancellationRequested();
            image = null;
            if (contentType == "image/jp2" &&
                (!OfficeJpeg2000Header.TryGetOpaqueDimensions(encodedBytes, out _, out int width, out int height) ||
                 !OfficeRasterImageDecoder.IsWithinPixelLimit(width, height, _maximumPixels))) {
                throw new NotSupportedException("The JPEG 2000 image dimensions exceed the supported raster limit or have an unsupported header.");
            }
            bool decoded = _codec != null && _codec.TryDecode(encodedBytes, contentType, out image);
            _token.ThrowIfCancellationRequested();
            if (!decoded || image == null ||
                !OfficeRasterImageDecoder.IsWithinPixelLimit(image.Width, image.Height, _maximumPixels)) {
                image = null;
                throw new NotSupportedException("Raster rendering cannot decode the image within the raster limit (" + (contentType ?? "unknown content type") + ").");
            }
            return true;
        }
    }
}
