using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.Epub.Image;

/// <summary>EPUB image-export entry points backed by OfficeIMO.Html.</summary>
public static partial class EpubImageExportExtensions {
    private static readonly HashSet<string> PackageOmissionDiagnosticCodes =
        new HashSet<string>(StringComparer.Ordinal) {
            "epub.archive.duplicate-path",
            "epub.archive.unsafe-path",
            "epub.chapter.encrypted",
            "epub.chapter.count-limit",
            "epub.chapter.invalid-encoding",
            "epub.chapter.invalid-xhtml",
            "epub.chapter.text-total-limit",
            "epub.chapter.raw-html-total-limit",
            "epub.chapter.size-limit",
            "epub.encryption.resource-missing",
            "epub.encryption.unsupported",
            "epub.manifest.duplicate-id",
            "epub.manifest.invalid-path",
            "epub.resource.count-limit",
            "epub.resource.encrypted",
            "epub.resource.missing",
            "epub.resource.size-limit",
            "epub.resource.total-size-limit",
            "epub.spine.manifest-id-missing",
            "epub.spine.fallback-missing",
            "epub.spine.fallback-cycle",
            "epub.spine.remote-resource",
            "epub.spine.resource-missing",
            "epub.spine.unsupported-media-type"
        };

    /// <summary>Exports selected EPUB chapters through the shared image result contract.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this EpubDocument source,
        OfficeImageExportFormat format,
        EpubImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        source.ExportImages(
            format,
            results.Add,
            options,
            cancellationToken);
        return results.AsReadOnly();
    }

    /// <summary>Streams selected EPUB chapter images without retaining earlier payloads.</summary>
    public static void ExportImages(
        this EpubDocument source,
        OfficeImageExportFormat format,
        OfficeImageExportConsumer consumer,
        EpubImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        EpubImageExportOptions effective =
            options?.CloneEpub() ?? new EpubImageExportOptions();
        IReadOnlyList<EpubChapter> chapters = SelectChapters(source, effective);
        IReadOnlyDictionary<string, EpubResource> resourcesByPath = BuildResourceIndex(source, cancellationToken);
        OfficeImageExportBatchProcessor.Run(
            effective,
            (accept, operationCancellationToken) => {
                for (int index = 0; index < chapters.Count; index++) {
                    operationCancellationToken.ThrowIfCancellationRequested();
                    EpubChapter chapter = chapters[index];
                    EpubChapterRenderPreparation preparation =
                        PrepareChapter(chapter, effective, resourcesByPath);
                    preparation.Document.ExportImages(
                        format,
                        result => accept(CompleteResult(
                            result,
                            source,
                            chapter,
                            preparation.Diagnostics,
                            effective)),
                        preparation.Options,
                        operationCancellationToken);
                }
            },
            consumer,
            cancellationToken,
            effective.Mode == HtmlRenderMode.Continuous ? chapters.Count : (int?)null);
    }

    /// <summary>Asynchronously exports selected EPUB chapters and resolves retained package resources.</summary>
    public static async Task<IReadOnlyList<OfficeImageExportResult>> ExportImagesAsync(
        this EpubDocument source,
        OfficeImageExportFormat format,
        EpubImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        await source.ExportImagesAsync(
            format,
            (result, token) => {
                results.Add(result);
                return Task.CompletedTask;
            },
            options,
            cancellationToken).ConfigureAwait(false);
        return results.AsReadOnly();
    }

    /// <summary>Asynchronously streams selected EPUB chapter images.</summary>
    public static async Task ExportImagesAsync(
        this EpubDocument source,
        OfficeImageExportFormat format,
        OfficeImageExportAsyncConsumer consumer,
        EpubImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        EpubImageExportOptions effective =
            options?.CloneEpub() ?? new EpubImageExportOptions();
        IReadOnlyList<EpubChapter> chapters = SelectChapters(source, effective);
        IReadOnlyDictionary<string, EpubResource> resourcesByPath = BuildResourceIndex(source, cancellationToken);
        await OfficeImageExportBatchProcessor.RunAsync(
            effective,
            async (accept, operationCancellationToken) => {
                foreach (EpubChapter chapter in chapters) {
                    operationCancellationToken.ThrowIfCancellationRequested();
                    EpubChapterRenderPreparation preparation =
                        PrepareChapter(chapter, effective, resourcesByPath);
                    await preparation.Document.ExportImagesAsync(
                        format,
                        async (result, token) => await accept(
                            CompleteResult(
                                result,
                                source,
                                chapter,
                                preparation.Diagnostics,
                                effective),
                            token).ConfigureAwait(false),
                        preparation.Options,
                        operationCancellationToken).ConfigureAwait(false);
                }
            },
            consumer,
            cancellationToken,
            effective.Mode == HtmlRenderMode.Continuous ? chapters.Count : (int?)null).ConfigureAwait(false);
    }

    /// <summary>Starts fluent image export for selected EPUB chapters.</summary>
    public static EpubImageExportBuilder ToImages(
        this EpubDocument source,
        EpubImageExportOptions? options = null) =>
        new EpubImageExportBuilder(source, options);

    private static IReadOnlyList<EpubChapter> SelectChapters(
        EpubDocument source,
        EpubImageExportOptions options) {
        if (options.ChapterIndex < 0) {
            throw new ArgumentOutOfRangeException(
                nameof(options.ChapterIndex));
        }
        if (options.ChapterCount.HasValue &&
            options.ChapterCount.Value < 1) {
            throw new ArgumentOutOfRangeException(
                nameof(options.ChapterCount));
        }
        if (options.ChapterIndex >= source.Chapters.Count) {
            if (source.Chapters.Count == 0) {
                throw new InvalidOperationException(
                    "The EPUB does not contain any extracted chapters.");
            }
            throw new ArgumentOutOfRangeException(
                nameof(options.ChapterIndex));
        }
        int available = source.Chapters.Count - options.ChapterIndex;
        int count = options.ChapterCount.HasValue
            ? Math.Min(options.ChapterCount.Value, available)
            : available;
        return source.Chapters
            .Skip(options.ChapterIndex)
            .Take(count)
            .ToArray();
    }

    private static EpubChapterRenderPreparation PrepareChapter(
        EpubChapter chapter,
        EpubImageExportOptions options,
        IReadOnlyDictionary<string, EpubResource> resourcesByPath) {
        EpubImageExportOptions effective = options.CloneEpub();
        effective.Policy = new OfficeImageExportPolicy();
        Uri baseUri = CreateChapterUri(chapter);
        effective.BaseUri = baseUri;
        ConfigureResources(resourcesByPath, effective);
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        string html;
        if (!string.IsNullOrWhiteSpace(chapter.Html)) {
            html = chapter.Html!;
        } else {
            html = CreatePlainTextChapter(chapter, options.IncludeChapterTitle);
            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                "EPUB_IMAGE_RAW_HTML_UNAVAILABLE",
                "Raw chapter HTML was not retained; the extracted chapter text was rendered instead.",
                chapter.Path,
                OfficeConversionLossKind.Approximation));
        }
        if (chapter.Encryption?.RequiresDecryption == true) {
            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                "EPUB_IMAGE_CHAPTER_ENCRYPTED",
                "The chapter declares unsupported encryption and may be incomplete.",
                chapter.Path,
                OfficeConversionLossKind.Omission));
        }
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            html,
            new HtmlConversionDocumentOptions {
                BaseUri = baseUri,
                UrlPolicy = effective.UrlPolicy.Clone(),
                ResourceUrlPolicy = (effective.ResourceUrlPolicy ?? effective.UrlPolicy).Clone(),
                UseBodyContentsOnly = false
            });
        return new EpubChapterRenderPreparation(
            document,
            effective,
            diagnostics.AsReadOnly());
    }

    private static IReadOnlyDictionary<string, EpubResource> BuildResourceIndex(
        EpubDocument source,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var resourcesByPath = new Dictionary<string, EpubResource>(StringComparer.Ordinal);
        foreach (EpubResource resource in source.Resources) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!resource.IsRemote && !resourcesByPath.ContainsKey(resource.Path)) {
                resourcesByPath.Add(resource.Path, resource);
            }
        }
        return resourcesByPath;
    }

    private static void ConfigureResources(
        IReadOnlyDictionary<string, EpubResource> resourcesByPath,
        EpubImageExportOptions options) {
        HtmlUrlPolicy fallbackResourceUrlPolicy =
            (options.ResourceUrlPolicy ??
             options.UrlPolicy ??
             HtmlUrlPolicy.CreateOfficeIMOProfile())
            .Clone();
        HtmlUrlPolicy resourcePolicy = fallbackResourceUrlPolicy.Clone();
        resourcePolicy.RestrictUrlSchemes = true;
        resourcePolicy.AllowedUrlSchemes.Add("epub");
        options.ResourceUrlPolicy = resourcePolicy;
        HtmlRenderSynchronousResourceResolver? synchronousFallback =
            options.SynchronousResourceResolver;
        options.SynchronousResourceResolver = (
            HtmlRenderResourceRequest request,
            CancellationToken cancellationToken,
            out HtmlResolvedResource? resolved) => {
            cancellationToken.ThrowIfCancellationRequested();
            EpubResource? resource = FindResource(
                resourcesByPath,
                request);
            byte[]? data = resource?.Data;
            if (data is { Length: > 0 }) {
                if (data.LongLength > options.MaxResourceBytes) {
                    throw new HtmlRenderResourceByteLimitException(
                        data.LongLength);
                }
                resolved = new HtmlResolvedResource(
                    data,
                    resource!.MediaType ?? "application/octet-stream");
                return true;
            }
            if (resource != null ||
                request.Uri.Scheme.Equals(
                    "epub",
                    StringComparison.OrdinalIgnoreCase)) {
                resolved = null;
                return true;
            }
            if (synchronousFallback != null &&
                HtmlUrlPolicyEvaluator.IsAllowed(
                    request.Uri.AbsoluteUri,
                    fallbackResourceUrlPolicy)) {
                return synchronousFallback(
                    request,
                    cancellationToken,
                    out resolved);
            }
            resolved = null;
            return false;
        };
        HtmlRenderResourceResolver? fallback = options.ResourceResolver;
        options.ResourceResolver = async (request, cancellationToken) => {
            cancellationToken.ThrowIfCancellationRequested();
            EpubResource? resource = FindResource(
                resourcesByPath,
                request);
            byte[]? data = resource?.Data;
            if (data is { Length: > 0 }) {
                if (data.LongLength > options.MaxResourceBytes) {
                    throw new HtmlRenderResourceByteLimitException(
                        data.LongLength);
                }
                return new HtmlResolvedResource(
                    data,
                    resource!.MediaType ?? "application/octet-stream");
            }
            if (resource != null ||
                request.Uri.Scheme.Equals(
                    "epub",
                    StringComparison.OrdinalIgnoreCase) ||
                fallback == null ||
                !HtmlUrlPolicyEvaluator.IsAllowed(
                    request.Uri.AbsoluteUri,
                    fallbackResourceUrlPolicy)) {
                return null;
            }
            return await fallback(request, cancellationToken)
                .ConfigureAwait(false);
        };
    }

    private static EpubResource? FindResource(
        IReadOnlyDictionary<string, EpubResource> resourcesByPath,
        HtmlRenderResourceRequest request) {
        // Only the package's virtual origin may select retained bytes. An HTML base
        // can change the request origin, but cannot change the package identity.
        if (!request.Uri.Scheme.Equals("epub", StringComparison.OrdinalIgnoreCase) ||
            !request.Uri.Host.Equals("document", StringComparison.OrdinalIgnoreCase) ||
            request.Uri.Port != -1 || request.Uri.UserInfo.Length != 0) return null;
        EpubReference reference = EpubReference.Resolve("package.opf", request.Uri.AbsolutePath);
        return reference.Kind == EpubReferenceKind.Container && reference.ContainerPath != null &&
            resourcesByPath.TryGetValue(reference.ContainerPath, out EpubResource? resource) ? resource : null;
    }

    private static OfficeImageExportResult CompleteResult(
        OfficeImageExportResult result,
        EpubDocument source,
        EpubChapter chapter,
        IReadOnlyList<OfficeImageExportDiagnostic> chapterDiagnostics,
        EpubImageExportOptions options) {
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        if (options.IncludePackageDiagnostics) {
            foreach (EpubDiagnostic diagnostic in source.Diagnostics) {
                OfficeImageExportDiagnosticSeverity severity =
                    diagnostic.Severity == EpubDiagnosticSeverity.Error
                        ? OfficeImageExportDiagnosticSeverity.Error
                        : diagnostic.Severity == EpubDiagnosticSeverity.Warning
                            ? OfficeImageExportDiagnosticSeverity.Warning
                            : OfficeImageExportDiagnosticSeverity.Info;
                diagnostics.Add(new OfficeImageExportDiagnostic(
                    severity,
                    "EPUB_IMAGE_" + NormalizeCode(diagnostic.Code),
                    diagnostic.Message,
                    diagnostic.Path,
                    severity == OfficeImageExportDiagnosticSeverity.Error
                        ? OfficeConversionLossKind.Failure
                        : severity == OfficeImageExportDiagnosticSeverity.Warning &&
                          IsPackageOmissionDiagnostic(diagnostic.Code)
                            ? OfficeConversionLossKind.Omission
                            : severity == OfficeImageExportDiagnosticSeverity.Warning
                                ? OfficeConversionLossKind.Approximation
                                : OfficeConversionLossKind.None));
            }
        }
        diagnostics.AddRange(chapterDiagnostics);
        diagnostics.AddRange(result.Diagnostics);
        string name = string.IsNullOrWhiteSpace(chapter.Title)
            ? "Chapter " + chapter.Order
            : chapter.Title!;
        if (options.Mode == HtmlRenderMode.Paged) {
            name += " - " + (result.Name ?? "Page");
        }
        return options.EnsureAccepted(new OfficeImageExportResult(
            result.Format,
            result.Width,
            result.Height,
            result.Bytes,
            name,
            chapter.Path,
            diagnostics));
    }

    private static bool IsPackageOmissionDiagnostic(string code) {
        return !string.IsNullOrWhiteSpace(code) &&
               PackageOmissionDiagnosticCodes.Contains(code);
    }

    private static string CreatePlainTextChapter(
        EpubChapter chapter,
        bool includeTitle) {
        var builder = new StringBuilder(
            "<!doctype html><html><head><meta charset=\"utf-8\"></head><body>");
        if (includeTitle && !string.IsNullOrWhiteSpace(chapter.Title)) {
            builder.Append("<h1>")
                .Append(WebUtility.HtmlEncode(chapter.Title))
                .Append("</h1>");
        }
        builder.Append("<div style=\"white-space:pre-wrap\">")
            .Append(WebUtility.HtmlEncode(chapter.Text))
            .Append("</div></body></html>");
        return builder.ToString();
    }

    private static Uri CreateChapterUri(EpubChapter chapter) {
        string path = NormalizePath(chapter.Path);
        // The HTML parser applies the retained markup's base element once.
        return new Uri("epub://document/" + EscapePath(path));
    }

    private static string EscapePath(string path) =>
        string.Join(
            "/",
            path.Split('/')
                .Where(segment => segment.Length > 0)
                .Select(Uri.EscapeDataString));

    private static string NormalizePath(string path) {
        var segments = new List<string>();
        foreach (string segment in path
                     .Replace('\\', '/')
                     .Split('/')) {
            if (segment.Length == 0 || segment == ".") continue;
            if (segment == "..") {
                if (segments.Count > 0) segments.RemoveAt(
                    segments.Count - 1);
                continue;
            }
            segments.Add(segment);
        }
        return string.Join("/", segments);
    }

    private static string NormalizeCode(string code) {
        if (string.IsNullOrWhiteSpace(code)) return "DIAGNOSTIC";
        var builder = new StringBuilder();
        bool underscore = false;
        foreach (char character in code) {
            char value = char.IsLetterOrDigit(character)
                ? char.ToUpperInvariant(character)
                : '_';
            if (value == '_' && underscore) continue;
            builder.Append(value);
            underscore = value == '_';
        }
        return builder.ToString().Trim('_');
    }

    private sealed class EpubChapterRenderPreparation {
        internal EpubChapterRenderPreparation(
            HtmlConversionDocument document,
            EpubImageExportOptions options,
            IReadOnlyList<OfficeImageExportDiagnostic> diagnostics) {
            Document = document;
            Options = options;
            Diagnostics = diagnostics;
        }

        internal HtmlConversionDocument Document { get; }
        internal EpubImageExportOptions Options { get; }
        internal IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics { get; }
    }
}
