using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PdfTextRun = OfficeIMO.Pdf.TextRun;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static partial class MarkdownPdfConverterExtensions {
    private static void RenderImageBlock(PdfCore.PdfDocument pdf, ImageBlock image, MarkdownPdfSaveOptions options, MarkdownPdfStyle visualTheme) {
        if (!TryReadImageBytes(image.Path, options, out byte[] bytes, out string sourceName, out string warningCode, out string warningMessage)) {
            AddWarning(options, warningCode, image.Path, warningMessage);
            RenderImagePlaceholder(pdf, image.PlainAlt ?? image.Alt ?? image.Path, visualTheme);
            return;
        }

        RenderImageFigure(
            pdf,
            bytes,
            sourceName,
            image.Width,
            image.Height,
            image.PlainAlt ?? image.Alt,
            image.Caption,
            NormalizeAbsoluteLink(image.LinkUrl),
            image.LinkTitle ?? image.PlainAlt ?? image.Alt,
            options,
            visualTheme);
    }

    private static bool TryRenderImageOnlyParagraph(PdfCore.PdfDocument pdf, InlineSequence inlines, MarkdownPdfSaveOptions options, MarkdownPdfStyle visualTheme) {
        if (inlines.Nodes.Count != 1) {
            return false;
        }

        if (inlines.Nodes[0] is ImageInline image) {
            RenderInlineImageFigure(pdf, image.Src, image.PlainAlt.Length == 0 ? image.Alt : image.PlainAlt, image.Title, linkUrl: null, linkTitle: null, options, visualTheme);
            return true;
        }

        if (inlines.Nodes[0] is ImageLinkInline imageLink) {
            RenderInlineImageFigure(pdf, imageLink.ImageUrl, imageLink.PlainAlt.Length == 0 ? imageLink.Alt : imageLink.PlainAlt, imageLink.Title, imageLink.LinkUrl, imageLink.LinkTitle, options, visualTheme);
            return true;
        }

        return false;
    }

    private static void RenderInlineImageFigure(
        PdfCore.PdfDocument pdf,
        string source,
        string? altText,
        string? title,
        string? linkUrl,
        string? linkTitle,
        MarkdownPdfSaveOptions options,
        MarkdownPdfStyle visualTheme) {
        if (!TryReadImageBytes(source, options, out byte[] bytes, out string sourceName, out string warningCode, out string warningMessage)) {
            AddWarning(options, warningCode, source, warningMessage);
            RenderImagePlaceholder(pdf, altText ?? source, visualTheme);
            return;
        }

        RenderImageFigure(
            pdf,
            bytes,
            sourceName,
            widthHint: null,
            heightHint: null,
            altText,
            caption: null,
            linkUri: NormalizeAbsoluteLink(linkUrl),
            linkContents: linkTitle ?? altText,
            options,
            visualTheme);
    }

    private static void RenderImageFigure(
        PdfCore.PdfDocument pdf,
        byte[] bytes,
        string sourceName,
        double? widthHint,
        double? heightHint,
        string? altText,
        string? caption,
        string? linkUri,
        string? linkContents,
        MarkdownPdfSaveOptions options,
        MarkdownPdfStyle visualTheme) {
        if (!PdfCore.PdfDocument.TryPrepareImageBytes(
                bytes,
                out byte[] preparedBytes,
                out OfficeImageInfo? info,
                out _,
                out string? unsupportedReason)) {
            AddWarning(options, "UnsupportedImage", sourceName, "The Markdown image bytes are not supported by the PDF image renderer. " + unsupportedReason);
            RenderImagePlaceholder(pdf, altText ?? sourceName, visualTheme);
            return;
        }

        double width = widthHint ?? GetImageWidthPoints(info, options);
        double height = heightHint ?? GetImageHeightPoints(info, width, options);
        MarkdownPdfFigureStyle figureStyle = visualTheme.FigureStyleSnapshot;
        PdfCore.PdfImageStyle imageStyle = CreateConverterImageStyle(figureStyle, altText);
        string? normalizedLinkContents = linkUri == null || string.IsNullOrWhiteSpace(linkContents) ? null : linkContents;

        pdf.Image(preparedBytes, width, height, align: null, clipPath: null, fit: null, spacingBefore: null, spacingAfter: null, style: imageStyle, linkUri: linkUri, linkContents: normalizedLinkContents);
        RenderFigureCaption(pdf, caption, figureStyle);
    }

    private static void RenderImagePlaceholder(PdfCore.PdfDocument pdf, string? label, MarkdownPdfStyle visualTheme) {
        if (string.IsNullOrWhiteSpace(label)) {
            label = "Image";
        }

        MarkdownPdfFigureStyle figureStyle = visualTheme.FigureStyleSnapshot;
        PdfCore.PanelStyle? panelStyle = figureStyle.PanelStyleSnapshot;
        Action<PdfCore.PdfParagraphBuilder> render = builder => builder
            .Italic(true)
            .Color(figureStyle.PlaceholderColorSnapshot)
            .Text("[Image unavailable: " + label + "]");

        if (panelStyle != null) {
            pdf.PanelParagraph(render, panelStyle, defaultColor: figureStyle.PlaceholderColorSnapshot);
            return;
        }

        pdf.Paragraph(render, defaultColor: figureStyle.PlaceholderColorSnapshot);
    }

    private static PdfCore.PdfImageStyle CreateConverterImageStyle(MarkdownPdfFigureStyle figureStyle, string? altText) {
        PdfCore.PdfImageStyle style = figureStyle.ImageStyleSnapshot;
        style.ScaleDownToFit = true;
        if (!string.IsNullOrWhiteSpace(altText)) {
            style.AlternativeText = altText!.Trim();
        }

        return style;
    }

    private static void RenderFigureCaption(PdfCore.PdfDocument pdf, string? caption, MarkdownPdfFigureStyle figureStyle) {
        if (string.IsNullOrWhiteSpace(caption)) {
            return;
        }

        pdf.Paragraph(builder => {
            builder.Italic(true)
                .FontSize(figureStyle.CaptionFontSizeSnapshot)
                .Color(figureStyle.CaptionColorSnapshot)
                .Text(caption!.Trim());
        }, align: figureStyle.CaptionAlignSnapshot, defaultColor: figureStyle.CaptionColorSnapshot, style: new PdfCore.PdfParagraphStyle { SpacingBefore = 2, SpacingAfter = 2 });
    }

    private static bool TryReadImageBytes(string path, MarkdownPdfSaveOptions options, out byte[] bytes, out string sourceName, out string warningCode, out string warningMessage) {
        bytes = Array.Empty<byte>();
        sourceName = string.Empty;
        warningCode = "UnsupportedImage";
        warningMessage = "Only resolvable local Markdown images or supported base64 data URI images are embedded in the Markdown PDF adapter.";

        if (!options.IncludeImages) {
            warningCode = "ImagesDisabled";
            warningMessage = "Markdown images are disabled by the selected PDF export profile.";
            return false;
        }

        if (IsDataUri(path)) {
            if (!options.ResourcePolicy.AllowDataUris) {
                warningCode = "DataUriImageDisabled";
                warningMessage = "Data URI images are disabled by the PDF resource policy.";
                return false;
            }
            return TryReadDataUriImageBytes(path, options, out bytes, out sourceName, out warningCode, out warningMessage);
        }

        if (TryCreateRemoteImageUri(path, out Uri? remoteUri)) {
            if (!options.ResourcePolicy.AllowRemoteResourceResolution) {
                warningCode = "RemoteImageDisabled";
                warningMessage = "Remote Markdown images are disabled by the PDF resource policy.";
                return false;
            }
            return TryReadRemoteImageBytes(remoteUri!, options, out bytes, out sourceName, out warningCode, out warningMessage);
        }

        if (!options.ResourcePolicy.AllowLocalFileAccess) {
            warningCode = "LocalImageDisabled";
            warningMessage = "Local Markdown images are disabled by the PDF resource policy.";
            return false;
        }

        if (!TryResolveImagePath(path, options, out string resolvedPath, out warningCode, out warningMessage)) {
            return false;
        }

        bytes = File.ReadAllBytes(resolvedPath);
        sourceName = resolvedPath;
        return true;
    }

    private static bool TryResolveImagePath(
        string path,
        MarkdownPdfSaveOptions options,
        out string resolvedPath,
        out string warningCode,
        out string warningMessage) {
        resolvedPath = string.Empty;
        warningCode = "UnsupportedImage";
        warningMessage = "Only resolvable local Markdown images or supported base64 data URI images are embedded in the Markdown PDF adapter.";

        if (string.IsNullOrWhiteSpace(path) || Uri.TryCreate(path, UriKind.Absolute, out Uri? uri) && !uri.IsFile) {
            return false;
        }

        string candidate = path;
        if (Uri.TryCreate(path, UriKind.Absolute, out Uri? fileUri) && fileUri.IsFile) {
            candidate = fileUri.LocalPath;
        } else if (!Path.IsPathRooted(candidate) && !string.IsNullOrWhiteSpace(options.BaseDirectory)) {
            candidate = Path.Combine(options.BaseDirectory!, candidate);
        }

        string fullPath;
        try {
            fullPath = Path.GetFullPath(candidate);
        } catch (Exception ex) when (ex is ArgumentException || ex is NotSupportedException || ex is PathTooLongException) {
            return false;
        }

        if (!File.Exists(fullPath)) {
            return false;
        }

        if (options.RestrictLocalImagesToBaseDirectory &&
            !string.IsNullOrWhiteSpace(options.BaseDirectory)) {
            try {
                if (!IsPathInsideDirectory(fullPath, options.BaseDirectory!)) {
                    warningCode = "LocalImageOutsideBaseDirectory";
                    warningMessage = "Local Markdown image paths must resolve inside MarkdownPdfSaveOptions.BaseDirectory.";
                    return false;
                }
            } catch (Exception ex) when (ex is ArgumentException || ex is IOException || ex is NotSupportedException || ex is PathTooLongException || ex is UnauthorizedAccessException) {
                warningCode = "LocalImageOutsideBaseDirectory";
                warningMessage = "Local Markdown image paths must resolve safely inside MarkdownPdfSaveOptions.BaseDirectory.";
                return false;
            }
        }

        resolvedPath = fullPath;
        return true;
    }

    private static bool IsPathInsideDirectory(string fullPath, string baseDirectory) {
        string normalizedBase = EnsureTrailingDirectorySeparator(ResolvePhysicalPath(Path.GetFullPath(baseDirectory)));
        string normalizedPath = ResolvePhysicalPath(Path.GetFullPath(fullPath));
        return normalizedPath.StartsWith(normalizedBase, GetPathComparison());
    }

    private static string ResolvePhysicalPath(string fullPath) {
        string normalized = Path.GetFullPath(fullPath);
        string root = Path.GetPathRoot(normalized) ?? string.Empty;
        if (root.Length == 0) throw new ArgumentException("A rooted resource path is required.", nameof(fullPath));

        string current = root;
        string[] segments = normalized.Substring(root.Length)
            .Split(new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries);
        for (int index = 0; index < segments.Length; index++) {
            string candidate = Path.Combine(current, segments[index]);
            bool isDirectory = Directory.Exists(candidate);
            bool isFile = File.Exists(candidate);
            if (!isDirectory && !isFile) throw new FileNotFoundException("A local PDF resource path component does not exist.", candidate);

#if NET8_0_OR_GREATER
            FileSystemInfo info = isDirectory ? (FileSystemInfo)new DirectoryInfo(candidate) : new FileInfo(candidate);
            FileSystemInfo? target = info.ResolveLinkTarget(returnFinalTarget: true);
            current = target == null ? candidate : target.FullName;
#else
            if ((File.GetAttributes(candidate) & FileAttributes.ReparsePoint) != 0) {
                throw new IOException("Symbolic links and reparse points are not accepted for restricted local PDF resources on this target framework.");
            }
            current = candidate;
#endif
        }

        return Path.GetFullPath(current);
    }

    private static string EnsureTrailingDirectorySeparator(string path) {
        if (path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) ||
            path.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal)) {
            return path;
        }

        return path + Path.DirectorySeparatorChar;
    }

    private static StringComparison GetPathComparison() =>
        Path.DirectorySeparatorChar == '\\' ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;

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
        var warning = new MarkdownPdfExportWarning(code, source, message);
        options.Warnings.Add(warning);
        options.Report.Add(warning.ToConversionWarning());
    }
}
