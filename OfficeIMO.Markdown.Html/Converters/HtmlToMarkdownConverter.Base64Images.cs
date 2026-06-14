using OfficeIMO.Html;

namespace OfficeIMO.Markdown.Html;

public sealed partial class HtmlToMarkdownConverter {
    private static bool TryApplyBase64ImageHandling(ref string source, ConversionContext context) {
        if (!TryParseBase64ImageDataUri(source, out string mimeType, out byte[] bytes)) {
            return true;
        }

        switch (context.Options.Base64Images) {
            case HtmlBase64ImageHandling.Include:
                return true;
            case HtmlBase64ImageHandling.Skip:
                source = string.Empty;
                return false;
            case HtmlBase64ImageHandling.SaveToFile:
                source = SaveBase64Image(bytes, mimeType, context);
                return source.Length > 0;
            default:
                throw new ArgumentOutOfRangeException(nameof(context.Options.Base64Images), context.Options.Base64Images, "Unknown base64 image handling mode.");
        }
    }

    private static bool TryParseBase64ImageDataUri(string? source, out string mimeType, out byte[] bytes) {
        mimeType = string.Empty;
        bytes = Array.Empty<byte>();
        if (!HtmlImageDataUri.TryParse(source, out var dataUri) || !dataUri.IsBase64) {
            return false;
        }

        if (!dataUri.TryDecodeBytes(out bytes)) {
            return false;
        }

        mimeType = dataUri.MediaType;
        return true;
    }

    private static string SaveBase64Image(byte[] bytes, string mimeType, ConversionContext context) {
        string? outputDirectory = context.Options.Base64ImageOutputDirectory;
        if (string.IsNullOrWhiteSpace(outputDirectory)) {
            throw new InvalidOperationException("Base64ImageOutputDirectory is required when Base64Images is SaveToFile.");
        }

        string directory = Path.GetFullPath(outputDirectory!);
        Directory.CreateDirectory(directory);

        int index = context.SavedBase64ImageCount++;
        string fileName = context.Options.Base64ImageFileNameGenerator?.Invoke(index, mimeType) ?? "image_" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
        fileName = SanitizeBase64ImageFileName(fileName, GetImageExtension(mimeType), index);

        string path = Path.Combine(directory, fileName);
        File.WriteAllBytes(path, bytes);
        return path;
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
