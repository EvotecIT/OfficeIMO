using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

namespace OfficeIMO.Workflows;

internal static class OfficeWorkflowHtmlResourceResolver {
    private const int BufferSize = 81920;

    internal static HtmlPdfSaveOptions CreateOptions(string inputPath) {
        if (string.IsNullOrWhiteSpace(inputPath)) {
            throw new ArgumentException("Input path cannot be empty.", nameof(inputPath));
        }

        string sourceDirectory = Path.GetDirectoryName(Path.GetFullPath(inputPath))
            ?? throw new ArgumentException("HTML input path must have a parent directory.", nameof(inputPath));
        string physicalRoot = OfficeWorkflowPathIdentity.ResolvePhysicalPath(sourceDirectory);
        HtmlUrlPolicy resourcePolicy = CreateResourcePolicy();

        var options = new HtmlPdfSaveOptions {
            ResourceUrlPolicy = resourcePolicy
        };
        options.ResourcePolicy.AllowLocalFileAccess = true;
        options.ResourceResolver = (request, cancellationToken) =>
            ResolveAsync(request, physicalRoot, options.MaxResourceBytes, cancellationToken);
        return options;
    }

    internal static HtmlUrlPolicy CreateResourcePolicy() {
        var resourcePolicy = new HtmlUrlPolicy {
            DisallowFileUrls = false,
            AllowMailtoUrls = false,
            AllowDataUrls = true,
            AllowProtocolRelativeUrls = false,
            RestrictUrlSchemes = true
        };
        resourcePolicy.AllowedUrlSchemes.Clear();
        resourcePolicy.AllowedUrlSchemes.Add(Uri.UriSchemeFile);
        resourcePolicy.AllowedUrlSchemes.Add("data");
        return resourcePolicy;
    }

    private static Task<HtmlResolvedResource?> ResolveAsync(
        HtmlRenderResourceRequest request,
        string physicalRoot,
        long maximumResourceBytes,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!request.Uri.IsFile || request.Kind == HtmlResourceKind.Hyperlink || request.Kind == HtmlResourceKind.Script) {
            return Task.FromResult<HtmlResolvedResource?>(null);
        }

        try {
            using FileStream source = OfficeWorkflowPathIdentity.OpenRegularFileForRead(
                request.Uri.LocalPath,
                physicalRoot,
                BufferSize);
            byte[] bytes = OfficeWorkflowInputReader.ReadAllBytes(
                source,
                Path.GetFileName(request.Uri.LocalPath),
                maximumResourceBytes,
                cancellationToken);
            if (bytes.Length == 0) return Task.FromResult<HtmlResolvedResource?>(null);
            return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(bytes, GetContentType(request)));
        } catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException or ArgumentException) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult<HtmlResolvedResource?>(null);
        }
    }

    private static string GetContentType(HtmlRenderResourceRequest request) {
        if (request.Kind == HtmlResourceKind.Stylesheet) return "text/css";
        if (request.Kind == HtmlResourceKind.Font) return GetFontContentType(request.Uri.LocalPath);
        if (request.Kind == HtmlResourceKind.Image) return GetImageContentType(request.Uri.LocalPath);
        return "application/octet-stream";
    }

    private static string GetImageContentType(string path) =>
        Path.GetExtension(path).ToLowerInvariant() switch {
            ".png" => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".webp" => "image/webp",
            ".svg" => "image/svg+xml",
            ".bmp" => "image/bmp",
            ".tif" or ".tiff" => "image/tiff",
            _ => "application/octet-stream"
        };

    private static string GetFontContentType(string path) =>
        Path.GetExtension(path).ToLowerInvariant() switch {
            ".woff" => "font/woff",
            ".woff2" => "font/woff2",
            ".ttf" => "font/ttf",
            ".otf" => "font/otf",
            _ => "application/octet-stream"
        };
}
