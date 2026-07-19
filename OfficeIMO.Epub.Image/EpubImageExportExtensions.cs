using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.Epub.Image;

/// <summary>EPUB image-export entry points backed by OfficeIMO.Html.</summary>
public static class EpubImageExportExtensions {
    private static readonly HashSet<string> PackageOmissionDiagnosticCodes =
        new HashSet<string>(StringComparer.Ordinal) {
            "epub.archive.duplicate-path",
            "epub.archive.unsafe-path",
            "epub.chapter.encrypted",
            "epub.chapter.invalid-xhtml",
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
            "epub.spine.remote-resource",
            "epub.spine.resource-missing"
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
        OfficeImageExportConsumer accept =
            OfficeImageExportBatchProcessor.CreateGuardedConsumer(
                effective,
                consumer,
                cancellationToken);
        for (int index = 0; index < chapters.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            EpubChapter chapter = chapters[index];
            EpubChapterRenderPreparation preparation =
                PrepareChapter(source, chapter, effective);
            preparation.Document.ExportImages(
                format,
                result => accept(CompleteResult(
                    result,
                    source,
                    chapter,
                    preparation.Diagnostics,
                    effective)),
                preparation.Options,
                cancellationToken);
        }
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
        OfficeImageExportAsyncConsumer accept =
            OfficeImageExportBatchProcessor.CreateGuardedAsyncConsumer(
                effective,
                consumer,
                cancellationToken);
        foreach (EpubChapter chapter in chapters) {
            cancellationToken.ThrowIfCancellationRequested();
            EpubChapterRenderPreparation preparation =
                PrepareChapter(source, chapter, effective);
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
                cancellationToken).ConfigureAwait(false);
        }
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
        EpubDocument source,
        EpubChapter chapter,
        EpubImageExportOptions options) {
        EpubImageExportOptions effective = options.CloneEpub();
        effective.Policy = new OfficeImageExportPolicy();
        Uri baseUri = CreateChapterUri(chapter);
        effective.BaseUri = baseUri;
        ConfigureResources(source, effective, baseUri);
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
                OfficeImageExportLossKind.Approximation));
        }
        if (chapter.Encryption?.RequiresDecryption == true) {
            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                "EPUB_IMAGE_CHAPTER_ENCRYPTED",
                "The chapter declares unsupported encryption and may be incomplete.",
                chapter.Path,
                OfficeImageExportLossKind.Omission));
        }
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            html,
            new HtmlConversionDocumentOptions {
                BaseUri = baseUri,
                UseBodyContentsOnly = false
            });
        return new EpubChapterRenderPreparation(
            document,
            effective,
            diagnostics.AsReadOnly());
    }

    private static void ConfigureResources(
        EpubDocument source,
        EpubImageExportOptions options,
        Uri baseUri) {
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
                source,
                request,
                baseUri);
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
                source,
                request,
                baseUri);
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
        EpubDocument source,
        HtmlRenderResourceRequest request,
        Uri baseUri) {
        string requestPath = NormalizePath(
            Uri.UnescapeDataString(request.Uri.AbsolutePath));
        foreach (EpubResource resource in source.Resources) {
            if (resource.IsRemote) continue;
            if (string.Equals(
                NormalizePath(resource.Path),
                requestPath,
                StringComparison.OrdinalIgnoreCase)) {
                return resource;
            }
            if (!string.IsNullOrWhiteSpace(resource.Href) &&
                Uri.TryCreate(baseUri, resource.Href, out Uri? resolved) &&
                string.Equals(
                    resolved.AbsoluteUri,
                    request.Uri.AbsoluteUri,
                    StringComparison.OrdinalIgnoreCase)) {
                return resource;
            }
        }
        return null;
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
                        ? OfficeImageExportLossKind.Failure
                        : severity == OfficeImageExportDiagnosticSeverity.Warning &&
                          IsPackageOmissionDiagnostic(diagnostic.Code)
                            ? OfficeImageExportLossKind.Omission
                            : severity == OfficeImageExportDiagnosticSeverity.Warning
                                ? OfficeImageExportLossKind.Approximation
                                : OfficeImageExportLossKind.None));
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
        Uri chapterUri = new Uri(
            "epub://document/" + EscapePath(path));
        if (!string.IsNullOrWhiteSpace(chapter.BaseHref) &&
            Uri.TryCreate(chapterUri, chapter.BaseHref, out Uri? resolved)) {
            return resolved;
        }
        return chapterUri;
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
