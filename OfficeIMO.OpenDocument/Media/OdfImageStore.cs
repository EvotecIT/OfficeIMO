namespace OfficeIMO.OpenDocument;

using OfficeIMO.Drawing;

internal static class OdfImageStore {
    internal static string Add(OdfDocument document, byte[] data, string fileName) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (data.Length == 0) throw new ArgumentException("Image data cannot be empty.", nameof(data));
        if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("Image file name cannot be empty.", nameof(fileName));

        string extension;
        if (!OdfImageFormats.TryNormalizeStoredExtension(fileName, out extension)) {
            if (!OfficeImageReader.TryIdentifyByContent(data, fileName, out OfficeImageInfo info) ||
                !OdfImageFormats.TryGetExtension(info.Format, out extension)) {
                throw new NotSupportedException("Supported image formats are PNG, JPEG, GIF, SVG, BMP, and WebP.");
            }
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
