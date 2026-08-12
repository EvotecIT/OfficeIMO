using OfficeIMO.Email;
using OfficeIMO.Mhtml;
using OfficeIMO.Reader.Html;

namespace OfficeIMO.Reader.Email;

internal static class MhtmlReaderAdapter {
    internal static IEnumerable<ReaderChunk> Read(
        string path,
        ReaderOptions? readerOptions,
        ReaderHtmlOptions? htmlOptions,
        CancellationToken cancellationToken) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("MHTML path cannot be empty.", nameof(path));
        if (!File.Exists(path)) throw new FileNotFoundException($"MHTML file '{path}' doesn't exist.", path);

        ReaderOptions effective = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(path, GetMaxInputBytes(effective));
        HtmlReaderAdapter.SourceMetadata source = HtmlReaderAdapter.BuildSourceMetadataFromPath(path, effective.ComputeHashes);
        MhtmlDocument archive = MhtmlDocument.Load(path, CreateMhtmlReaderOptions(effective), cancellationToken: cancellationToken);
        return HtmlReaderAdapter.ReadContent(archive.Html, source, effective,
            PrepareHtmlOptions(htmlOptions, archive), cancellationToken).ToArray();
    }

    internal static IEnumerable<ReaderChunk> Read(
        Stream stream,
        string? sourceName,
        ReaderOptions? readerOptions,
        ReaderHtmlOptions? htmlOptions,
        CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("MHTML stream must be readable.", nameof(stream));

        ReaderOptions effective = readerOptions ?? new ReaderOptions();
        long maximumInputBytes = GetMaxInputBytes(effective);
        string logicalSourceName = string.IsNullOrWhiteSpace(sourceName) ? "document.mhtml" : sourceName!.Trim();
        var source = new HtmlReaderAdapter.SourceMetadata {
            Path = logicalSourceName,
            SourceId = HtmlReaderAdapter.BuildSourceId(logicalSourceName)
        };
        Stream parseStream = ReaderInputLimits.EnsureSeekableReadStream(stream, maximumInputBytes,
            cancellationToken, out bool ownsParseStream);
        try {
            HtmlReaderAdapter.UpdateSourceMetadataFromSeekableStream(source, parseStream, effective.ComputeHashes);
            MhtmlDocument archive = MhtmlDocument.Load(parseStream, CreateMhtmlReaderOptions(effective),
                cancellationToken: cancellationToken);
            return HtmlReaderAdapter.ReadContent(archive.Html, source, effective,
                PrepareHtmlOptions(htmlOptions, archive), cancellationToken).ToArray();
        } finally {
            if (ownsParseStream) parseStream.Dispose();
        }
    }

    internal static OfficeDocumentReadResult ReadDocument(
        string path,
        ReaderOptions? readerOptions,
        ReaderHtmlOptions? htmlOptions,
        CancellationToken cancellationToken) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("MHTML path cannot be empty.", nameof(path));
        if (!File.Exists(path)) throw new FileNotFoundException($"MHTML file '{path}' doesn't exist.", path);

        ReaderOptions effective = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(path, GetMaxInputBytes(effective));
        HtmlReaderAdapter.SourceMetadata source = HtmlReaderAdapter.BuildSourceMetadataFromPath(path, effective.ComputeHashes);
        MhtmlDocument archive = MhtmlDocument.Load(path, CreateMhtmlReaderOptions(effective), cancellationToken: cancellationToken);
        return ProjectDocument(archive, source, effective, htmlOptions, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadDocument(
        Stream stream,
        string? sourceName,
        ReaderOptions? readerOptions,
        ReaderHtmlOptions? htmlOptions,
        CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("MHTML stream must be readable.", nameof(stream));

        ReaderOptions effective = readerOptions ?? new ReaderOptions();
        long maximumInputBytes = GetMaxInputBytes(effective);
        string logicalSourceName = string.IsNullOrWhiteSpace(sourceName) ? "document.mhtml" : sourceName!.Trim();
        var source = new HtmlReaderAdapter.SourceMetadata {
            Path = logicalSourceName,
            SourceId = HtmlReaderAdapter.BuildSourceId(logicalSourceName)
        };
        Stream parseStream = ReaderInputLimits.EnsureSeekableReadStream(stream, maximumInputBytes,
            cancellationToken, out bool ownsParseStream);
        try {
            HtmlReaderAdapter.UpdateSourceMetadataFromSeekableStream(source, parseStream, effective.ComputeHashes);
            MhtmlDocument archive = MhtmlDocument.Load(parseStream, CreateMhtmlReaderOptions(effective),
                cancellationToken: cancellationToken);
            return ProjectDocument(archive, source, effective, htmlOptions, cancellationToken);
        } finally {
            if (ownsParseStream) parseStream.Dispose();
        }
    }

    private static OfficeDocumentReadResult ProjectDocument(
        MhtmlDocument archive,
        HtmlReaderAdapter.SourceMetadata source,
        ReaderOptions readerOptions,
        ReaderHtmlOptions? htmlOptions,
        CancellationToken cancellationToken) {
        OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(
            archive.Html, source, readerOptions, PrepareHtmlOptions(htmlOptions, archive), cancellationToken);
        MergeResources(result, archive, source.Path);
        if (string.IsNullOrWhiteSpace(result.Source.Title)) result.Source.Title = archive.Subject;
        result.CapabilitiesUsed = result.CapabilitiesUsed
            .Concat(new[] { "officeimo.reader.mhtml", "officeimo.mhtml" })
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        result.Diagnostics = result.Diagnostics.Concat(MapDiagnostics(archive, source.Path)).ToArray();
        return result;
    }

    private static EmailReaderOptions CreateMhtmlReaderOptions(ReaderOptions options) {
        return new EmailReaderOptions(maxInputBytes: GetMaxInputBytes(options));
    }

    private static long GetMaxInputBytes(ReaderOptions options) =>
        options.MaxInputBytes ?? OfficeDocumentReaderBuilderMhtmlExtensions.DefaultMaxInputBytes;

    private static ReaderHtmlOptions PrepareHtmlOptions(ReaderHtmlOptions? source, MhtmlDocument archive) {
        ReaderHtmlOptions options = ReaderHtmlOptionsCloner.CloneOrDefault(source);
        var projection = options.HtmlToMarkdownOptions ?? OfficeIMO.Markdown.Html.HtmlToMarkdownOptions.CreateOfficeIMOProfile();
        projection.BaseUri ??= archive.BaseUri;
        if (projection.UrlPolicy.RestrictUrlSchemes) {
            projection.UrlPolicy.AllowedUrlSchemes.Add("cid");
            projection.UrlPolicy.AllowedUrlSchemes.Add(archive.BaseUri.Scheme);
        }
        options.HtmlToMarkdownOptions = projection;
        var conversion = options.ConversionOptions?.Clone() ?? OfficeIMO.Html.HtmlConversionDocumentOptions.CreateUntrustedProfile();
        conversion.BaseUri ??= archive.BaseUri;
        if (conversion.ResourceUrlPolicy.RestrictUrlSchemes) {
            conversion.ResourceUrlPolicy.AllowedUrlSchemes.Add("cid");
            conversion.ResourceUrlPolicy.AllowedUrlSchemes.Add(archive.BaseUri.Scheme);
        }
        options.ConversionOptions = conversion;
        return options;
    }

    private static void MergeResources(OfficeDocumentReadResult result, MhtmlDocument archive, string? path) {
        var assets = result.Assets.ToList();
        var visuals = result.Visuals.ToList();
        int resourceIndex = 0;
        foreach (MhtmlResource resource in archive.Resources) {
            byte[] content = resource.Content;
            OfficeDocumentAsset[] matches = assets
                .Where(asset => MatchesResource(asset.SourceObjectId, resource, archive.BaseUri))
                .ToArray();
            if (matches.Length == 0) {
                string id = "mhtml-resource-" + resourceIndex.ToString("D4", CultureInfo.InvariantCulture);
                string? extension = ResolveExtension(resource);
                string? fileName = ResolveFileName(resource);
                var asset = new OfficeDocumentAsset {
                    Id = id,
                    Kind = ResolveResourceKind(resource.ContentType),
                    MediaType = resource.ContentType,
                    Extension = extension,
                    FileName = fileName ?? (extension == null ? null : OfficeDocumentAssetNaming.BuildFileName(id, extension)),
                    SourceObjectId = ResolveSourceId(resource),
                    Location = new ReaderLocation { Path = path, SourceBlockKind = "mhtml-resource", BlockAnchor = id }
                };
                assets.Add(asset);
                matches = new[] { asset };
            }

            foreach (OfficeDocumentAsset asset in matches) {
                asset.MediaType = resource.ContentType;
                asset.Extension ??= ResolveExtension(resource);
                asset.FileName ??= ResolveFileName(resource);
                asset.LengthBytes = content.LongLength;
                asset.PayloadHash = HtmlReaderAdapter.ComputeHtmlHash(content);
                asset.PayloadBytes = (byte[])content.Clone();
                asset.SourceObjectId = ResolveSourceId(resource);
                ReaderVisual? visual = visuals.FirstOrDefault(candidate =>
                    string.Equals(candidate.Location?.BlockAnchor, asset.Location.BlockAnchor, StringComparison.Ordinal));
                if (visual != null) {
                    visual.PayloadHash = asset.PayloadHash;
                    visual.SourceName = asset.SourceObjectId;
                    visual.MimeType = resource.ContentType;
                }
            }
            resourceIndex++;
        }
        result.Assets = assets;
        result.Visuals = visuals;
    }

    private static bool MatchesResource(string? source, MhtmlResource resource, Uri baseUri) {
        if (string.IsNullOrWhiteSpace(source)) return false;
        string sourceValue = source!;
        if (!string.IsNullOrWhiteSpace(resource.ContentId) && sourceValue.StartsWith("cid:", StringComparison.OrdinalIgnoreCase)) {
            string contentId = Uri.UnescapeDataString(sourceValue.Substring("cid:".Length)).Trim().Trim('<', '>');
            if (string.Equals(contentId, resource.ContentId, StringComparison.OrdinalIgnoreCase)) return true;
        }
        string? contentLocation = resource.ContentLocation;
        if (string.IsNullOrWhiteSpace(contentLocation)) return false;
        if (string.Equals(sourceValue, contentLocation, StringComparison.OrdinalIgnoreCase)) return true;
        return Uri.TryCreate(baseUri, contentLocation, out Uri? resourceUri) &&
               string.Equals(sourceValue, resourceUri.AbsoluteUri, StringComparison.OrdinalIgnoreCase);
    }

    private static string ResolveResourceKind(string contentType) {
        if (contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) return "image";
        if (string.Equals(contentType, "text/css", StringComparison.OrdinalIgnoreCase)) return "stylesheet";
        if (contentType.IndexOf("javascript", StringComparison.OrdinalIgnoreCase) >= 0) return "script";
        if (contentType.StartsWith("font/", StringComparison.OrdinalIgnoreCase)) return "font";
        return "resource";
    }

    private static string? ResolveExtension(MhtmlResource resource) {
        string? fileName = ResolveFileName(resource);
        string extension = string.IsNullOrWhiteSpace(fileName) ? string.Empty : Path.GetExtension(fileName);
        if (!string.IsNullOrWhiteSpace(extension)) return extension;
        return resource.ContentType.ToLowerInvariant() switch {
            "image/png" => ".png",
            "image/jpeg" => ".jpg",
            "image/gif" => ".gif",
            "image/svg+xml" => ".svg",
            "image/webp" => ".webp",
            "text/css" => ".css",
            "text/javascript" or "application/javascript" => ".js",
            "font/woff" => ".woff",
            "font/woff2" => ".woff2",
            _ => null
        };
    }

    private static string? ResolveFileName(MhtmlResource resource) {
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

    private static string? ResolveSourceId(MhtmlResource resource) {
        if (!string.IsNullOrWhiteSpace(resource.ContentId)) return "cid:" + resource.ContentId;
        if (!string.IsNullOrWhiteSpace(resource.ContentLocation)) return resource.ContentLocation;
        return resource.FileName;
    }

    private static IEnumerable<OfficeDocumentDiagnostic> MapDiagnostics(MhtmlDocument archive, string? path) {
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
                Source = "OfficeIMO.Reader.Email.Mhtml",
                IsRecoverable = diagnostic.Severity != EmailDiagnosticSeverity.Error,
                Location = new ReaderLocation { Path = path }
            };
        }
    }

}
