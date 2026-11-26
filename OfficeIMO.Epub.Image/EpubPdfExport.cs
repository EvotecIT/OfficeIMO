using System.Diagnostics;
using System.Globalization;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Epub.Image;

internal static class EpubPdfExport {
    internal static bool AllowsPackageResources(HtmlPdfExportOptions? options) =>
        options?.ResourcePolicy.AllowEmbeddedPackageResources != false &&
        (options?.ResourceUrlPolicy is not { RestrictUrlSchemes: true } policy || policy.AllowedUrlSchemes.Contains("epub"));

    internal static OfficeConversionReport Write(EpubDocument book, Stream destination, HtmlPdfExportOptions? options,
        bool failOnLoss, CancellationToken cancellationToken, int? importedHtmlNodeLimit = null) {
        HtmlPdfExportOptions rendering = options?.ClonePdf() ?? new HtmlPdfExportOptions();
        if (importedHtmlNodeLimit is int maximum) {
            rendering.MaxHtmlNodes = rendering.MaxHtmlNodes == HtmlConversionLimits.UntrustedDefaults.MaxHtmlNodes
                ? maximum : Math.Min(rendering.MaxHtmlNodes, maximum);
        }
        return HtmlRenderEngine.ExecuteWithDeadline(rendering, cancellationToken,
            token => WriteCore(book, destination, rendering, failOnLoss, token, importedHtmlNodeLimit.HasValue));
    }

    private static OfficeConversionReport WriteCore(EpubDocument book, Stream destination, HtmlPdfExportOptions rendering,
        bool failOnLoss, CancellationToken cancellationToken, bool importedPublication) {
        cancellationToken.ThrowIfCancellationRequested();
        if (book.Chapters.Count == 0) throw new InvalidDataException("The EPUB does not contain any chapters to render.");
        var timer = Stopwatch.StartNew();
        failOnLoss |= rendering.FailOnLoss || rendering.FidelityPolicy == HtmlRenderFidelityPolicy.FailOnLoss;
        // The package is the resource authority. The renderer must not invoke callbacks or open host files.
        rendering.ResourceResolver = null;
        rendering.SynchronousResourceResolver = null;
        bool allowEmbedded = AllowsPackageResources(rendering);
        var resourcePolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile();
        if (allowEmbedded) resourcePolicy.AllowedUrlSchemes.Add("epub");
        resourcePolicy.AllowDataUrls &= rendering.ResourcePolicy.AllowDataUris;
        if (rendering.ResourceUrlPolicy != null) resourcePolicy = HtmlUrlPolicy.Intersect(resourcePolicy, rendering.ResourceUrlPolicy);
        resourcePolicy.ResolvedUrlTransform = null;
        rendering.ResourceUrlPolicy = resourcePolicy;
        rendering.UrlPolicy.AllowedUrlSchemes.Add("epub");
        rendering.InteractiveFormControls = false;
        rendering.SharePublicationResources();
        var resources = new EpubChapterResources(book, cancellationToken);
        var diagnostics = new List<OfficeDiagnostic>();
        foreach (EpubDiagnostic diagnostic in book.Diagnostics) {
            OfficeDiagnosticSeverity severity = diagnostic.Severity switch {
                EpubDiagnosticSeverity.Error => OfficeDiagnosticSeverity.Error,
                EpubDiagnosticSeverity.Warning => OfficeDiagnosticSeverity.Warning,
                _ => OfficeDiagnosticSeverity.Information
            };
            diagnostics.Add(new OfficeDiagnostic(diagnostic.Code, severity, diagnostic.Message, diagnostic.Path, "OfficeIMO.Epub",
                severity == OfficeDiagnosticSeverity.Error ? OfficeConversionLossKind.Failure
                : severity == OfficeDiagnosticSeverity.Warning ? OfficeConversionLossKind.Approximation : OfficeConversionLossKind.None));
        }
        var chaptersByPath = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (EpubChapter chapter in book.Chapters)
            if (!chaptersByPath.ContainsKey(chapter.Path)) chaptersByPath.Add(chapter.Path, chapter.Order);
        rendering.ResolveSectionLink = ResolveSectionLink;
        var outline = new Dictionary<int, List<HtmlPdfOutlineTarget>>();
        int outlineOrder = 0;
        IndexOutline(book.TableOfContents, 1);

        var reports = new List<(string? Chapter, PdfConversionReport Report)>();
        int pageCount = HtmlPdfRenderedConverter.WriteSections(destination, ReadChapters(), rendering,
            book.Title, book.Creator, book.Language, (chapter, report) => reports.Add((chapter, report)), cancellationToken);
        foreach ((string? chapter, PdfConversionReport report) in reports) {
            foreach (PdfConversionWarning warning in report.Warnings) {
                OfficeDiagnosticSeverity severity = warning.Severity switch {
                    PdfConversionWarningSeverity.Error => OfficeDiagnosticSeverity.Error,
                    PdfConversionWarningSeverity.Warning => OfficeDiagnosticSeverity.Warning,
                    _ => OfficeDiagnosticSeverity.Information
                };
                string location = chapter == null ? warning.Source : chapter + ": " + warning.Source;
                diagnostics.Add(new OfficeDiagnostic(warning.Code, severity, warning.Message, location, warning.Converter, warning.LossKind));
            }
        }

        var result = new OfficeConversionReport("EPUB", "PDF", itemsWritten: pageCount, elapsed: timer.Elapsed, diagnostics: diagnostics);
        if (failOnLoss) result.ThrowIfLoss();
        return result;

        string? ResolveSectionLink(Uri uri) {
            if (!uri.Scheme.Equals("epub", StringComparison.OrdinalIgnoreCase)) return null;
            if (!uri.Host.Equals("document", StringComparison.OrdinalIgnoreCase) || uri.Port != -1 || uri.UserInfo.Length != 0)
                return string.Empty;
            EpubReference link = EpubReference.Resolve("package.opf", uri.PathAndQuery + uri.Fragment);
            return link.Kind == EpubReferenceKind.Container && link.ContainerPath != null &&
                chaptersByPath.TryGetValue(link.ContainerPath, out int order)
                ? DestinationPrefix(order) + (string.IsNullOrEmpty(link.Fragment) ? HtmlRenderOptions.PublicationStartAnchor :
                    HtmlRenderOptions.PublicationFragmentPrefix + link.Fragment)
                : string.Empty;
        }

        (int Order, string? Fragment)? IndexOutline(IReadOnlyList<EpubNavigationItem> items, int level) {
            (int Order, string? Fragment)? firstTarget = null;
            foreach (EpubNavigationItem item in items) {
                cancellationToken.ThrowIfCancellationRequested();
                bool labeled = !string.IsNullOrWhiteSpace(item.Label);
                (int Order, string? Fragment)? target = null;
                if (!item.IsRemote && item.Target != null && chaptersByPath.TryGetValue(item.Target, out int chapterOrder))
                    target = (chapterOrder, item.Fragment);
                else if (item.Href != null) {
                    diagnostics.Add(new OfficeDiagnostic("EPUB_PDF_OUTLINE_TARGET_UNAVAILABLE", OfficeDiagnosticSeverity.Warning,
                        "A navigation entry targets content outside the rendered spine.", item.Href, "OfficeIMO.Epub", OfficeConversionLossKind.Omission));
                }
                bool group = labeled && item.Href == null;
                bool hasOutline = labeled && (target.HasValue || group);
                int order = hasOutline ? outlineOrder++ : outlineOrder;
                int childLevel = hasOutline ? Math.Min(level + 1, PdfPageCanvas.MaximumOutlineLevel) : level;
                var childTarget = IndexOutline(item.Children, childLevel);
                if (group) target ??= childTarget;
                if (hasOutline && target is { } resolved) {
                    if (!outline.TryGetValue(resolved.Order, out List<HtmlPdfOutlineTarget>? entries))
                        outline[resolved.Order] = entries = new List<HtmlPdfOutlineTarget>();
                    entries.Add(new HtmlPdfOutlineTarget(item.Label, resolved.Fragment, level, order));
                }
                firstTarget ??= target ?? childTarget;
            }
            return firstTarget;
        }

        IEnumerable<HtmlPdfSection> ReadChapters() {
            foreach (EpubChapter chapter in book.Chapters) {
                cancellationToken.ThrowIfCancellationRequested();
                var chapterOptions = rendering.ClonePdf();
                chapterOptions.NamedDestinationPrefix = DestinationPrefix(chapter.Order);
                IReadOnlyList<HtmlPdfOutlineTarget> entries = outline.TryGetValue(chapter.Order, out var chapterOutline)
                    ? chapterOutline : Array.Empty<HtmlPdfOutlineTarget>();
                chapterOptions.AdditionalNamedDestinations = entries.Select(entry => entry.Fragment)
                    .OfType<string>().ToArray();
                if (string.IsNullOrWhiteSpace(chapter.Html)) {
                    diagnostics.Add(new OfficeDiagnostic("EPUB_PDF_RAW_HTML_UNAVAILABLE", OfficeDiagnosticSeverity.Warning,
                        "Extracted text was rendered because chapter HTML was unavailable.", chapter.Path, "OfficeIMO.Epub", OfficeConversionLossKind.Approximation));
                }
                HtmlConversionDocument document = resources.PrepareDocument(chapter, chapterOptions, includeChapterTitle: true, cancellationToken,
                    allowEmbeddedResources: allowEmbedded, maximumHtmlNodes: importedPublication ? chapterOptions.MaxHtmlNodes : null);
                var profile = HtmlRenderIntentProfile.PrintPaged;
                if (chapter.IsFixedLayout) {
                    if (EpubFixedLayoutViewport.TryRead(document, out double width, out double height)) {
                        chapterOptions.ViewportWidth = width;
                        chapterOptions.ViewportHeight = height;
                        chapterOptions.Margins = HtmlRenderMargins.All(0D);
                        profile = HtmlRenderIntentProfile.ScreenViewport;
                    } else {
                        diagnostics.Add(new OfficeDiagnostic("EPUB_PDF_FIXED_LAYOUT_VIEWPORT_UNAVAILABLE", OfficeDiagnosticSeverity.Warning,
                            "A pre-paginated chapter without a usable viewport was rendered as reflowable HTML.",
                            chapter.Path, "OfficeIMO.Epub", OfficeConversionLossKind.Approximation));
                    }
                }
                var encoder = profile == HtmlRenderIntentProfile.ScreenViewport ? HtmlRenderEncoder.DisplayList : HtmlRenderEncoder.Pdf;
                var request = HtmlRenderRequest.Create(profile, encoder, chapterOptions);
                yield return new HtmlPdfSection(chapter.Path, document, request, entries);
            }
        }
    }
    private static string DestinationPrefix(int order) => order.ToString(CultureInfo.InvariantCulture) + ":";
}
