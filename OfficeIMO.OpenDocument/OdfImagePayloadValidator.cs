using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

internal static class OdfImagePayloadValidator {
    internal static bool TryResolvePreservedFileName(
        byte[]? bytes,
        string? fileName,
        out string storedFileName,
        OdfImageValidationBudget? budget = null) {
        storedFileName = string.Empty;
        if (bytes == null || budget != null && !budget.TryConsume(bytes, fileName)) return false;
        if (!OfficeImageReader.TryValidateContent(bytes, fileName, out OfficeImageInfo info)) {
            return false;
        }
        if (!OdfImageFormats.TryGetExtension(info.Format, out string extension)) return false;
        storedFileName = "image" + extension;
        return true;
    }
}
