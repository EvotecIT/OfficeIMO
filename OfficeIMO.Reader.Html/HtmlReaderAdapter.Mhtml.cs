using OfficeIMO.Email;
using OfficeIMO.Html;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Reader;
using System.Linq;

namespace OfficeIMO.Reader.Html;

internal static partial class HtmlReaderAdapter {
    private static bool IsMhtmlSource(string? sourceName) {
        if (string.IsNullOrWhiteSpace(sourceName)) return false;
        string extension = Path.GetExtension(sourceName);
        return string.Equals(extension, ".mht", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(extension, ".mhtml", StringComparison.OrdinalIgnoreCase);
    }

    private static MhtmlDocument LoadMhtml(Stream stream, ReaderOptions options,
        CancellationToken cancellationToken) =>
        MhtmlDocument.Load(stream, CreateMhtmlReaderOptions(options), cancellationToken: cancellationToken);

    private static EmailReaderOptions CreateMhtmlReaderOptions(ReaderOptions options) {
        long maxInputBytes = options.MaxInputBytes ?? EmailReaderOptions.Default.MaxInputBytes;
        return new EmailReaderOptions(maxInputBytes: maxInputBytes);
    }

    private static ReaderHtmlOptions PrepareMhtmlHtmlOptions(ReaderHtmlOptions? source, MhtmlDocument archive) {
        ReaderHtmlOptions options = ReaderHtmlOptionsCloner.CloneOrDefault(source);
        HtmlToMarkdownOptions projection = options.HtmlToMarkdownOptions ?? HtmlToMarkdownOptions.CreateOfficeIMOProfile();
        projection.BaseUri ??= archive.BaseUri;
        if (projection.UrlPolicy.RestrictUrlSchemes) {
            projection.UrlPolicy.AllowedUrlSchemes.Add("cid");
            projection.UrlPolicy.AllowedUrlSchemes.Add(archive.BaseUri.Scheme);
        }
        options.HtmlToMarkdownOptions = projection;
        HtmlConversionDocumentOptions conversion = options.ConversionOptions?.Clone()
            ?? HtmlConversionDocumentOptions.CreateUntrustedProfile();
        conversion.BaseUri ??= archive.BaseUri;
        if (conversion.ResourceUrlPolicy.RestrictUrlSchemes) {
            conversion.ResourceUrlPolicy.AllowedUrlSchemes.Add("cid");
            conversion.ResourceUrlPolicy.AllowedUrlSchemes.Add(archive.BaseUri.Scheme);
        }
        options.ConversionOptions = conversion;
        return options;
    }

    private static void MergeMhtmlResources(HtmlProjection projection, MhtmlDocument archive, string? path) {
        int resourceIndex = 0;
        foreach (MhtmlResource resource in archive.Resources) {
            byte[] content = resource.Content;
            OfficeDocumentAsset[] matches = projection.Assets
                .Where(asset => MatchesMhtmlResource(asset.SourceObjectId, resource, archive.BaseUri))
                .ToArray();
            if (matches.Length == 0) {
                string id = "mhtml-resource-" + resourceIndex.ToString("D4", CultureInfo.InvariantCulture);
                string? extension = ResolveMhtmlExtension(resource);
                string? fileName = ResolveMhtmlFileName(resource);
                matches = new[] {
                    new OfficeDocumentAsset {
                        Id = id,
                        Kind = ResolveMhtmlResourceKind(resource.ContentType),
                        MediaType = resource.ContentType,
                        Extension = extension,
                        FileName = fileName ?? (extension == null ? null : OfficeDocumentAssetNaming.BuildFileName(id, extension)),
                        SourceObjectId = ResolveMhtmlSourceId(resource),
                        Location = BuildHtmlLocation(path, null, "mhtml-resource", id)
                    }
                };
                projection.Assets.Add(matches[0]);
            }

            foreach (OfficeDocumentAsset asset in matches) {
                asset.MediaType = resource.ContentType;
                asset.Extension ??= ResolveMhtmlExtension(resource);
                asset.FileName ??= ResolveMhtmlFileName(resource);
                asset.LengthBytes = content.LongLength;
                asset.PayloadHash = ComputeHtmlHash(content);
                asset.PayloadBytes = (byte[])content.Clone();
                asset.SourceObjectId = ResolveMhtmlSourceId(resource);
                ReaderVisual? visual = projection.Visuals.FirstOrDefault(candidate =>
                    string.Equals(candidate.Location?.BlockAnchor, asset.Location.BlockAnchor,
                        StringComparison.Ordinal));
                if (visual != null) {
                    visual.PayloadHash = asset.PayloadHash;
                    visual.SourceName = asset.SourceObjectId;
                    visual.MimeType = resource.ContentType;
                }
            }
            resourceIndex++;
        }
    }

    private static bool MatchesMhtmlResource(string? source, MhtmlResource resource, Uri baseUri) {
        if (string.IsNullOrWhiteSpace(source)) return false;
        string sourceValue = source!;
        if (!string.IsNullOrWhiteSpace(resource.ContentId) && sourceValue.StartsWith("cid:", StringComparison.OrdinalIgnoreCase)) {
            string contentId = Uri.UnescapeDataString(sourceValue.Substring("cid:".Length)).Trim().Trim('<', '>');
            if (string.Equals(contentId, resource.ContentId, StringComparison.OrdinalIgnoreCase)) return true;
        }
        if (string.IsNullOrWhiteSpace(resource.ContentLocation)) return false;
        if (string.Equals(sourceValue, resource.ContentLocation, StringComparison.OrdinalIgnoreCase)) return true;
        return Uri.TryCreate(baseUri, resource.ContentLocation, out Uri? resourceUri) &&
               string.Equals(sourceValue, resourceUri.AbsoluteUri, StringComparison.OrdinalIgnoreCase);
    }

    private static string ResolveMhtmlResourceKind(string contentType) {
        if (contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) return "image";
        if (string.Equals(contentType, "text/css", StringComparison.OrdinalIgnoreCase)) return "stylesheet";
        if (contentType.IndexOf("javascript", StringComparison.OrdinalIgnoreCase) >= 0) return "script";
        if (contentType.StartsWith("font/", StringComparison.OrdinalIgnoreCase)) return "font";
        return "resource";
    }

    private static string? ResolveMhtmlExtension(MhtmlResource resource) {
        string? fileName = ResolveMhtmlFileName(resource);
        string extension = string.IsNullOrWhiteSpace(fileName) ? string.Empty : Path.GetExtension(fileName);
        if (!string.IsNullOrWhiteSpace(extension)) return extension;
        switch (resource.ContentType.ToLowerInvariant()) {
            case "image/png": return ".png";
            case "image/jpeg": return ".jpg";
            case "image/gif": return ".gif";
            case "image/svg+xml": return ".svg";
            case "image/webp": return ".webp";
            case "text/css": return ".css";
            case "text/javascript":
            case "application/javascript": return ".js";
            case "font/woff": return ".woff";
            case "font/woff2": return ".woff2";
            default: return null;
        }
    }

    private static string? ResolveMhtmlFileName(MhtmlResource resource) {
        if (!string.IsNullOrWhiteSpace(resource.FileName)) return Path.GetFileName(resource.FileName);
        if (string.IsNullOrWhiteSpace(resource.ContentLocation)) return null;
        if (Uri.TryCreate(resource.ContentLocation, UriKind.Absolute, out Uri? uri)) {
            return Path.GetFileName(Uri.UnescapeDataString(uri.AbsolutePath));
        }
        string location = resource.ContentLocation!;
        int suffix = location.IndexOfAny(new[] { '?', '#' });
        if (suffix >= 0) location = location.Substring(0, suffix);
        return Path.GetFileName(location.Replace('/', Path.DirectorySeparatorChar));
    }

    private static string? ResolveMhtmlSourceId(MhtmlResource resource) {
        if (!string.IsNullOrWhiteSpace(resource.ContentId)) return "cid:" + resource.ContentId;
        if (!string.IsNullOrWhiteSpace(resource.ContentLocation)) return resource.ContentLocation;
        return resource.FileName;
    }

    private static IEnumerable<OfficeDocumentDiagnostic> MapMhtmlDiagnostics(MhtmlDocument archive, string? path) {
        foreach (EmailDiagnostic diagnostic in archive.MimeDiagnostics) {
            yield return new OfficeDocumentDiagnostic {
                Severity = diagnostic.Severity switch {
                    EmailDiagnosticSeverity.Information => OfficeDocumentDiagnosticSeverity.Information,
                    EmailDiagnosticSeverity.Error => OfficeDocumentDiagnosticSeverity.Error,
                    _ => OfficeDocumentDiagnosticSeverity.Warning
                },
                Category = OfficeDocumentDiagnosticCategory.Parsing,
                Code = diagnostic.Code,
                Message = diagnostic.Message,
                Source = "OfficeIMO.Html.Mhtml",
                IsRecoverable = diagnostic.Severity != EmailDiagnosticSeverity.Error,
                Location = new ReaderLocation { Path = path }
            };
        }
    }
}
