namespace OfficeIMO.OpenDocument;

using OfficeIMO.Drawing;

internal static class OdfImageStore {
    internal static string Add(OdfDocument document, byte[] data, string fileName) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (data.Length == 0) throw new ArgumentException("Image data cannot be empty.", nameof(data));
        if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("Image file name cannot be empty.", nameof(fileName));

        if (!OfficeImageReader.TryValidateContent(data, fileName, out OfficeImageInfo info)) {
            throw new ArgumentException("Image data must contain a complete supported image payload.", nameof(data));
        }
        if (!OdfImageFormats.TryGetExtension(info.Format, out string extension)) {
            throw new NotSupportedException("Supported image formats are PNG, JPEG, GIF, SVG, BMP, TIFF, and WebP.");
        }
        if (OdfImageFormats.TryNormalizeStoredExtension(fileName, out string declaredExtension) &&
            !string.Equals(declaredExtension, extension, StringComparison.Ordinal)) {
            throw new ArgumentException("Image file extension does not match the detected payload format.", nameof(fileName));
        }
        OdfImageFormats.TryGetMediaType(extension, out string mediaType);
        string hash;
        using (var algorithm = System.Security.Cryptography.SHA256.Create()) {
            hash = BitConverter.ToString(algorithm.ComputeHash(data)).Replace("-", string.Empty).ToLowerInvariant();
        }
        string path = "Pictures/" + hash.Substring(0, 24) + extension;
        if (!document.Package.ContainsEntry(path)) {
            document.Package.AddOrReplaceEntry(path, data, mediaType);
        }
        return path;
    }

}
