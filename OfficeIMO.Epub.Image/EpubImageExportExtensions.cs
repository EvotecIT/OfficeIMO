using OfficeIMO.Html.Pdf;
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
    public static IReadOnlyList<OfficeImageExportResult> ToImages(
        this EpubDocument source,
        OfficeImageExportFormat format,
        EpubImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        source.ToImages(
            format,
            results.Add,
            options,
            cancellationToken);
        return results.AsReadOnly();
    }

    /// <summary>Streams selected EPUB chapter images without retaining earlier payloads.</summary>
    public static void ToImages(
        this EpubDocument source,
        OfficeImageExportFormat format,
        OfficeImageExportConsumer consumer,
        EpubImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(consumer);
        EpubImageExportOptions effective =
            options?.CloneEpub() ?? new EpubImageExportOptions();
        IReadOnlyList<EpubChapter> chapters = SelectChapters(source, effective);
        var resources = new EpubChapterResources(source, cancellationToken);
        OfficeImageExportBatchProcessor.Run(
            effective,
            (accept, operationCancellationToken) => {
                for (int index = 0; index < chapters.Count; index++) {
                    operationCancellationToken.ThrowIfCancellationRequested();
                    EpubChapter chapter = chapters[index];
                    EpubChapterRenderPreparation preparation =
                        PrepareChapter(chapter, effective, resources, operationCancellationToken);
                    preparation.Document.ToImages(
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
    public static async Task<IReadOnlyList<OfficeImageExportResult>> ToImagesAsync(
        this EpubDocument source,
        OfficeImageExportFormat format,
        EpubImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        await source.ToImagesAsync(
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
    public static async Task ToImagesAsync(
        this EpubDocument source,
        OfficeImageExportFormat format,
        OfficeImageExportAsyncConsumer consumer,
        EpubImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(consumer);
        EpubImageExportOptions effective =
            options?.CloneEpub() ?? new EpubImageExportOptions();
        IReadOnlyList<EpubChapter> chapters = SelectChapters(source, effective);
        var resources = new EpubChapterResources(source, cancellationToken);
        await OfficeImageExportBatchProcessor.RunAsync(
            effective,
            async (accept, operationCancellationToken) => {
                foreach (EpubChapter chapter in chapters) {
                    operationCancellationToken.ThrowIfCancellationRequested();
                    EpubChapterRenderPreparation preparation =
                        PrepareChapter(chapter, effective, resources, operationCancellationToken);
                    await preparation.Document.ToImagesAsync(
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

    private static IReadOnlyList<EpubChapter> SelectChapters(
        EpubDocument source,
        EpubImageExportOptions options) {
        if (options.PageSelector != null)
            return options.PageSelector.Resolve(source.Chapters.Count, options.MaximumOutputCount).Select(number => source.Chapters[number - 1]).ToArray();
        ArgumentOutOfRangeException.ThrowIfNegative(options.ChapterIndex, nameof(options.ChapterIndex));
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
        EpubChapterResources resources,
        CancellationToken cancellationToken) {
        EpubImageExportOptions effective = options.CloneEpub();
        // The shared selector was consumed as chapter selection, before per-chapter pagination.
        effective.PageSelector = null;
        effective.Policy = new OfficeImageExportPolicy();
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        if (string.IsNullOrWhiteSpace(chapter.Html)) {
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
        HtmlConversionDocument document = resources.PrepareDocument(chapter, effective, options.IncludeChapterTitle, cancellationToken);
        return new EpubChapterRenderPreparation(document, effective, diagnostics.AsReadOnly());
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
