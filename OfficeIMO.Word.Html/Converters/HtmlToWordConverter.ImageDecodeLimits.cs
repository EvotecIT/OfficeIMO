using OfficeIMO.Drawing;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        // Inspect before Word normalizes WebP or AVIF into an allocated RGBA frame and PNG.
        private static void EnsureDecodedImageWithinLimits(Stream stream, HtmlToWordOptions options) {
            if (!options.MaxDecodedImagePixels.HasValue) return;
            if (OfficeImageReader.TryIdentifyByContent(stream, null, out OfficeImageInfo info)
                && info.Format is OfficeImageFormat.Webp or OfficeImageFormat.Avif
                && (long)info.Width * info.Height > options.MaxDecodedImagePixels.Value) {
                throw new HtmlResourceLimitException("Image dimensions exceed the configured decoded-pixel limit.");
            }
        }
    }
}
