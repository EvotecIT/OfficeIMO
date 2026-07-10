namespace OfficeIMO.OpenDocument;

internal static class OdfImageStore {
    internal static string Add(OdfDocument document, byte[] data, string fileName) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (data.Length == 0) throw new ArgumentException("Image data cannot be empty.", nameof(data));
        if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("Image file name cannot be empty.", nameof(fileName));

        string extension = NormalizeExtension(Path.GetExtension(fileName), data);
        string mediaType = MediaType(extension);
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

    private static string NormalizeExtension(string extension, byte[] data) {
        string value = extension.ToLowerInvariant();
        if (value == ".png" || value == ".jpg" || value == ".jpeg" || value == ".gif" || value == ".svg" || value == ".bmp" || value == ".webp") {
            return value == ".jpeg" ? ".jpg" : value;
        }
        if (data.Length >= 8 && data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47) return ".png";
        if (data.Length >= 3 && data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF) return ".jpg";
        if (data.Length >= 6 && data[0] == (byte)'G' && data[1] == (byte)'I' && data[2] == (byte)'F') return ".gif";
        throw new NotSupportedException("Supported image formats are PNG, JPEG, GIF, SVG, BMP, and WebP.");
    }

    private static string MediaType(string extension) {
        switch (extension) {
            case ".png": return "image/png";
            case ".jpg": return "image/jpeg";
            case ".gif": return "image/gif";
            case ".svg": return "image/svg+xml";
            case ".bmp": return "image/bmp";
            case ".webp": return "image/webp";
            default: return "application/octet-stream";
        }
    }
}
