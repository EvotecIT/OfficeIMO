using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PdfTextRun = OfficeIMO.Pdf.TextRun;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static partial class MarkdownPdfConverterExtensions {
    private static void RenderImageBlock(PdfCore.PdfDocument pdf, ImageBlock image, MarkdownPdfSaveOptions options) {
        if (!options.IncludeLocalImages) {
            RenderImagePlaceholder(pdf, image);
            return;
        }

        if (!TryReadImageBytes(image.Path, options, out byte[] bytes, out string sourceName, out string warningCode, out string warningMessage)) {
            AddWarning(options, warningCode, image.Path, warningMessage);
            RenderImagePlaceholder(pdf, image);
            return;
        }

        OfficeImageInfo? info = OfficeImageReader.TryIdentify(bytes, sourceName, out OfficeImageInfo? detected)
            ? detected
            : null;
        double width = image.Width ?? GetImageWidthPoints(info, options);
        double height = image.Height ?? GetImageHeightPoints(info, width, options);
        string? linkUri = NormalizeAbsoluteLink(image.LinkUrl);
        pdf.Image(bytes, width, height, PdfCore.PdfAlign.Left, spacingBefore: 4, spacingAfter: 6, linkUri: linkUri, linkContents: linkUri == null ? null : image.PlainAlt ?? image.Alt);

        if (!string.IsNullOrWhiteSpace(image.Caption)) {
            pdf.Paragraph(builder => builder.Italic(image.Caption!), style: new PdfCore.PdfParagraphStyle { SpacingAfter = 8 });
        }
    }

    private static void RenderImagePlaceholder(PdfCore.PdfDocument pdf, ImageBlock image) {
        string label = image.PlainAlt ?? image.Alt ?? image.Path;
        if (string.IsNullOrWhiteSpace(label)) {
            label = "Image";
        }

        pdf.Paragraph(builder => builder.Italic("[Image: " + label + "]"));
    }


    private static bool TryReadImageBytes(string path, MarkdownPdfSaveOptions options, out byte[] bytes, out string sourceName, out string warningCode, out string warningMessage) {
        bytes = Array.Empty<byte>();
        sourceName = string.Empty;
        warningCode = "UnsupportedImage";
        warningMessage = "Only resolvable local Markdown images or supported base64 data URI images are embedded in the Markdown PDF adapter.";

        if (IsDataUri(path)) {
            return TryReadDataUriImageBytes(path, options, out bytes, out sourceName, out warningCode, out warningMessage);
        }

        if (TryCreateRemoteImageUri(path, out Uri? remoteUri)) {
            return TryReadRemoteImageBytes(remoteUri!, options, out bytes, out sourceName, out warningCode, out warningMessage);
        }

        string? resolvedPath = ResolveImagePath(path, options.BaseDirectory);
        if (resolvedPath == null) {
            return false;
        }

        bytes = File.ReadAllBytes(resolvedPath);
        sourceName = resolvedPath;
        return true;
    }

    private static string? ResolveImagePath(string path, string? baseDirectory) {
        if (string.IsNullOrWhiteSpace(path) || Uri.TryCreate(path, UriKind.Absolute, out Uri? uri) && !uri.IsFile) {
            return null;
        }

        string candidate = path;
        if (Uri.TryCreate(path, UriKind.Absolute, out Uri? fileUri) && fileUri.IsFile) {
            candidate = fileUri.LocalPath;
        } else if (!Path.IsPathRooted(candidate) && !string.IsNullOrWhiteSpace(baseDirectory)) {
            candidate = Path.Combine(baseDirectory!, candidate);
        }

        return File.Exists(candidate) ? Path.GetFullPath(candidate) : null;
    }

    private static bool TryReadRemoteImageBytes(Uri uri, MarkdownPdfSaveOptions options, out byte[] bytes, out string sourceName, out string warningCode, out string warningMessage) {
        bytes = Array.Empty<byte>();
        sourceName = uri.ToString();
        warningCode = "UnsupportedImage";
        warningMessage = "Remote Markdown images require MarkdownPdfSaveOptions.RemoteImageResolver so callers can choose their own download, cache, and trust policy.";

        if (options.RemoteImageResolver == null) {
            return false;
        }

        byte[]? resolvedBytes;
        try {
            resolvedBytes = options.RemoteImageResolver(uri);
        } catch (Exception ex) when (ex is not OutOfMemoryException) {
            warningCode = "RemoteImageResolverFailed";
            warningMessage = "The configured Markdown remote image resolver failed: " + ex.Message;
            return false;
        }

        if (resolvedBytes == null || resolvedBytes.Length == 0) {
            warningMessage = "The configured Markdown remote image resolver did not return image bytes.";
            return false;
        }

        if (resolvedBytes.Length > options.MaximumRemoteImageBytes) {
            warningCode = "ImageTooLarge";
            warningMessage = "The resolved Markdown remote image exceeds the configured maximum byte length.";
            return false;
        }

        bytes = resolvedBytes;
        return true;
    }

    private static bool TryReadDataUriImageBytes(string path, MarkdownPdfSaveOptions options, out byte[] bytes, out string sourceName, out string warningCode, out string warningMessage) {
        bytes = Array.Empty<byte>();
        sourceName = string.Empty;
        warningCode = "UnsupportedImage";
        warningMessage = "The Markdown data URI image is not a supported base64 PNG or JPEG image.";

        if (!options.IncludeDataUriImages) {
            warningMessage = "Data URI images are disabled for this Markdown PDF export.";
            return false;
        }

        int commaIndex = path.IndexOf(',');
        if (commaIndex < 0) {
            return false;
        }

        string metadata = path.Substring("data:".Length, commaIndex - "data:".Length);
        string payload = path.Substring(commaIndex + 1);
        string[] parts = metadata.Split(';');
        string mediaType = parts.Length == 0 ? string.Empty : parts[0].Trim().ToLowerInvariant();
        if (mediaType != "image/png" && mediaType != "image/jpeg" && mediaType != "image/jpg") {
            return false;
        }

        bool isBase64 = false;
        for (int i = 1; i < parts.Length; i++) {
            if (string.Equals(parts[i].Trim(), "base64", StringComparison.OrdinalIgnoreCase)) {
                isBase64 = true;
                break;
            }
        }

        if (!isBase64) {
            return false;
        }

        string compactPayload = RemoveAsciiWhitespace(payload);
        long estimatedBytes = compactPayload.Length * 3L / 4L;
        if (estimatedBytes > options.MaximumDataUriImageBytes + 2L) {
            warningCode = "ImageTooLarge";
            warningMessage = "The decoded Markdown data URI image exceeds the configured maximum byte length.";
            return false;
        }

        try {
            bytes = Convert.FromBase64String(compactPayload);
        } catch (FormatException) {
            return false;
        }

        if (bytes.Length > options.MaximumDataUriImageBytes) {
            bytes = Array.Empty<byte>();
            warningCode = "ImageTooLarge";
            warningMessage = "The decoded Markdown data URI image exceeds the configured maximum byte length.";
            return false;
        }

        sourceName = mediaType == "image/png" ? "data-uri.png" : "data-uri.jpg";
        return true;
    }

    private static bool IsDataUri(string path) {
        return path.StartsWith("data:", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryCreateRemoteImageUri(string path, out Uri? uri) {
        uri = null;
        if (!Uri.TryCreate(path, UriKind.Absolute, out Uri? parsed)) {
            return false;
        }

        if (parsed.Scheme != Uri.UriSchemeHttp && parsed.Scheme != Uri.UriSchemeHttps) {
            return false;
        }

        uri = parsed;
        return true;
    }

    private static string RemoveAsciiWhitespace(string value) {
        StringBuilder? builder = null;
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n') {
                if (builder == null) {
                    builder = new StringBuilder(value.Length);
                    builder.Append(value, 0, i);
                }

                continue;
            }

            builder?.Append(ch);
        }

        return builder == null ? value : builder.ToString();
    }

    private static double GetImageWidthPoints(OfficeImageInfo? info, MarkdownPdfSaveOptions options) {
        if (info == null || info.Width <= 0) {
            return options.DefaultImageWidth;
        }

        return info.Width * 72D / info.DpiX;
    }

    private static double GetImageHeightPoints(OfficeImageInfo? info, double width, MarkdownPdfSaveOptions options) {
        if (info == null || info.Height <= 0 || info.Width <= 0) {
            return options.DefaultImageHeight;
        }

        return width * info.Height / info.Width;
    }


    private static void AddWarning(MarkdownPdfSaveOptions options, string code, string source, string message) {
        options.Warnings.Add(new MarkdownPdfExportWarning(code, source, message));
    }
}
