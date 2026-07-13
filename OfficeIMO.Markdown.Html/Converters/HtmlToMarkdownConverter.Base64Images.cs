using OfficeIMO.Html;

namespace OfficeIMO.Markdown.Html;

internal sealed partial class HtmlToMarkdownConverter {
    private static bool TryApplyBase64ImageHandling(ref string source, ConversionContext context) {
        string originalSource = source;
        if (!HtmlImageDataUri.TryParse(source, out var dataUri) || !dataUri.IsBase64) {
            return true;
        }

        switch (context.Options.Base64Images) {
            case HtmlBase64ImageHandling.Include:
                return true;
            case HtmlBase64ImageHandling.Skip:
                source = string.Empty;
                return false;
            case HtmlBase64ImageHandling.SaveToFile:
                if (context.SavedBase64ImagesBySource.TryGetValue(originalSource, out string? savedPath)) {
                    source = savedPath;
                    return source.Length > 0;
                }

                if (!dataUri.TryDecodeBytes(out byte[] bytes)) {
                    source = string.Empty;
                    return false;
                }

                source = EncodeMarkdownUrlWhitespace(SaveBase64Image(bytes, dataUri.MediaType, context));
                if (source.Length > 0) {
                    context.SavedBase64ImagesBySource[originalSource] = source;
                }

                return source.Length > 0;
            default:
                throw new ArgumentOutOfRangeException(nameof(context.Options.Base64Images), context.Options.Base64Images, "Unknown base64 image handling mode.");
        }
    }

    private static bool IsBase64ImageDataUri(string? source) {
        return HtmlImageDataUri.TryParse(source, out var dataUri) && dataUri.IsBase64;
    }

    private static string SaveBase64Image(byte[] bytes, string mimeType, ConversionContext context) {
        string? outputDirectory = context.Options.Base64ImageOutputDirectory;
        if (string.IsNullOrWhiteSpace(outputDirectory)) {
            throw new InvalidOperationException("Base64ImageOutputDirectory is required when Base64Images is SaveToFile.");
        }

        string directory = Path.GetFullPath(outputDirectory!);
        Directory.CreateDirectory(directory);

        int index = context.SavedBase64ImageCount++;
        string requestedFileName = context.Options.Base64ImageFileNameGenerator?.Invoke(index, mimeType) ?? "image_" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
        string fileName = SanitizeBase64ImageFileName(requestedFileName, GetImageExtension(mimeType), index);
        string path = Path.Combine(directory, fileName);
        return WriteBase64ImageFile(path, bytes);
    }

    private static string WriteBase64ImageFile(string path, byte[] bytes) {
        string directory = Path.GetDirectoryName(path) ?? string.Empty;
        string name = Path.GetFileNameWithoutExtension(path);
        string extension = Path.GetExtension(path);
        for (int attempt = 0; attempt < int.MaxValue; attempt++) {
            string candidate = attempt == 0
                ? path
                : Path.Combine(directory, name + "-" + attempt.ToString(System.Globalization.CultureInfo.InvariantCulture) + extension);
            try {
                using var stream = new FileStream(candidate, FileMode.CreateNew, FileAccess.Write);
                stream.Write(bytes, 0, bytes.Length);
                return candidate;
            } catch (IOException) when (File.Exists(candidate)) {
                continue;
            }
        }

        throw new IOException("Unable to create a unique file name for the decoded base64 image.");
    }

    private static string SanitizeBase64ImageFileName(string? fileName, string extension, int index) {
        string value = string.IsNullOrWhiteSpace(fileName)
            ? "image_" + index.ToString(System.Globalization.CultureInfo.InvariantCulture)
            : fileName!.Trim();

        foreach (char invalid in Path.GetInvalidFileNameChars()) {
            value = value.Replace(invalid, '_');
        }

        if (value.Length == 0) {
            value = "image_" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        return Path.HasExtension(value)
            ? value
            : value + extension;
    }

    private static string GetImageExtension(string mimeType) {
        return HtmlImageDataUri.GetFileExtension(mimeType);
    }
}
