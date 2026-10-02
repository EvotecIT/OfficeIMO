using OfficeIMO.Drawing;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        // Inspect before Word normalizes WebP into an allocated RGBA frame and PNG.
        private static void EnsureDecodedImageWithinLimits(Stream stream, HtmlToWordOptions options) {
            if (!options.MaxDecodedImagePixels.HasValue) return;
            if (OfficeImageReader.TryIdentifyByContent(stream, null, out OfficeImageInfo info)
                && info.Format == OfficeImageFormat.Webp
                && (long)info.Width * info.Height > options.MaxDecodedImagePixels.Value) {
                throw new HtmlResourceLimitException("Image dimensions exceed the configured decoded-pixel limit.");
            }
        }
    }
}
