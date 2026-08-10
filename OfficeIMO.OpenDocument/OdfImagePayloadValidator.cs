using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

internal static class OdfImagePayloadValidator {
    internal static bool CanPreserve(byte[]? bytes, string? fileName) {
        if (!OfficeImageReader.TryIdentifyByContent(bytes, fileName, out OfficeImageInfo info)) {
            return false;
        }

        // OpenDocument packages can preserve general WebP bitstreams without decoding them.
        // Other formats continue through the strict decoder-backed ingestion contract.
        return info.Format == OfficeImageFormat.Webp ||
               OfficeImageReader.TryValidateContent(bytes, fileName, out _);
    }
}
