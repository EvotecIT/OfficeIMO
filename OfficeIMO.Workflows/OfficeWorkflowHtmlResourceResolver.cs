using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

namespace OfficeIMO.Workflows;

internal static class OfficeWorkflowHtmlResourceResolver {
    private const int BufferSize = 81920;
    internal const int MaximumReferencedResourceCount = 256;
    internal const int MaximumStylesheetImportDepth = 16;
    private static readonly HashSet<string> SupportedDependencyExtensions = new(StringComparer.OrdinalIgnoreCase) {
        ".css", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".bmp", ".tif", ".tiff",
        ".woff", ".woff2", ".ttf", ".otf"
    };

    internal static bool IsSupportedDependency(string path) =>
        SupportedDependencyExtensions.Contains(Path.GetExtension(path));

    internal static HtmlPdfSaveOptions CreateOptions(
        string inputPath,
        long maximumResourceBytes,
        IReadOnlyDictionary<string, byte[]>? resourceSnapshots = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) {
            throw new ArgumentException("Input path cannot be empty.", nameof(inputPath));
        }
        if (maximumResourceBytes < 0L) {
            throw new ArgumentOutOfRangeException(nameof(maximumResourceBytes));
        }

        string sourceDirectory = Path.GetDirectoryName(Path.GetFullPath(inputPath))
            ?? throw new ArgumentException("HTML input path must have a parent directory.", nameof(inputPath));
        string physicalRoot = OfficeWorkflowPathIdentity.ResolvePhysicalPath(sourceDirectory);
        HtmlUrlPolicy resourcePolicy = CreateResourcePolicy();

        long rendererResourceBudget = Math.Max(1L, maximumResourceBytes);
        var options = new HtmlPdfSaveOptions {
            ResourceUrlPolicy = resourcePolicy,
            MaxResourceBytes = rendererResourceBudget,
            MaxTotalResourceBytes = rendererResourceBudget,
            MaxResourceCount = MaximumReferencedResourceCount,
            MaxResourceRequests = MaximumReferencedResourceCount * 2,
            MaxStylesheetImportDepth = MaximumStylesheetImportDepth
        };
        options.ResourcePolicy.AllowLocalFileAccess = true;
        options.ResourceResolver = (request, cancellationToken) =>
            ResolveAsync(request, physicalRoot, maximumResourceBytes, resourceSnapshots, cancellationToken);
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

    internal static HtmlResourcePipelineOptions CreatePdfResourcePipelineOptions(Uri? baseUri = null) {
        var renderOptions = new HtmlPdfSaveOptions();
        return new HtmlResourcePipelineOptions {
            BaseUri = baseUri,
            ResourceUrlPolicy = CreateResourcePolicy(),
            MediaContext = renderOptions.MediaContext,
            MediaWidth = renderOptions.PageWidth,
            MediaHeight = renderOptions.PageHeight
        };
    }

    private static Task<HtmlResolvedResource?> ResolveAsync(
        HtmlRenderResourceRequest request,
        string physicalRoot,
        long maximumResourceBytes,
        IReadOnlyDictionary<string, byte[]>? resourceSnapshots,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!request.Uri.IsFile || request.Kind == HtmlResourceKind.Hyperlink || request.Kind == HtmlResourceKind.Script) {
            return Task.FromResult<HtmlResolvedResource?>(null);
        }
        if (maximumResourceBytes <= 0L) {
            return Task.FromResult<HtmlResolvedResource?>(null);
        }

        if (resourceSnapshots is not null) {
            string fullPath = Path.GetFullPath(request.Uri.LocalPath);
            if (!resourceSnapshots.TryGetValue(fullPath, out byte[]? snapshot) ||
                snapshot.Length == 0 ||
                snapshot.LongLength > maximumResourceBytes) {
                return Task.FromResult<HtmlResolvedResource?>(null);
            }
            return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(snapshot, GetContentType(request)));
        }

        try {
            using FileStream source = OfficeWorkflowPathIdentity.OpenRegularFileForRead(
                request.Uri.LocalPath,
                physicalRoot,
                BufferSize);
            if (source.Length - source.Position > maximumResourceBytes) {
                return Task.FromResult<HtmlResolvedResource?>(null);
            }
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
