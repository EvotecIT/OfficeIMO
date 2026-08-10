using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

internal static class OdfImagePayloadValidator {
    internal static bool TryResolvePreservedFileName(
        byte[]? bytes,
        string? fileName,
        out string storedFileName) {
        storedFileName = string.Empty;
        if (!OfficeImageReader.TryIdentifyByContent(bytes, fileName, out OfficeImageInfo info)) {
            return false;
        }

        // OpenDocument packages can preserve general WebP bitstreams without decoding them.
        // Other formats continue through the strict decoder-backed ingestion contract.
        if (info.Format != OfficeImageFormat.Webp &&
            !OfficeImageReader.TryValidateContent(bytes, fileName, out _)) {
            return false;
        }
        if (!OdfImageFormats.TryGetExtension(info.Format, out string extension)) return false;
        storedFileName = "image" + extension;
        return true;
    }
}
