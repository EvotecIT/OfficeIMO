using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal sealed class HtmlPdfSection {
    internal HtmlPdfSection(string source, HtmlConversionDocument document, HtmlRenderRequest request, IReadOnlyList<HtmlPdfOutlineTarget> outline) {
        Source = source;
        Document = document;
        Request = request;
        Outline = outline;
    }

    internal string Source { get; }
    internal HtmlConversionDocument Document { get; }
    internal HtmlRenderRequest Request { get; }
    internal IReadOnlyList<HtmlPdfOutlineTarget> Outline { get; }
}

internal readonly struct HtmlPdfOutlineTarget {
    internal HtmlPdfOutlineTarget(string title, string? fragment, int level, int order) {
        Title = title;
        Fragment = fragment;
        Level = level;
        Order = order;
    }

    internal string Title { get; }
    internal string? Fragment { get; }
    internal int Level { get; }
    internal int Order { get; }
}

internal static partial class HtmlPdfRenderedConverter {
    // The PDF writer consumes pages from the current section and keeps its own bounded page-content store.
    internal static int WriteSections(Stream destination, IEnumerable<HtmlPdfSection> sections, HtmlPdfExportOptions options,
        string? title, string? author, string? language, Action<string?, PdfCore.PdfConversionReport> report,
        CancellationToken cancellationToken) {
        var bookOptions = options.PdfOptions.Clone();
        bookOptions.UseContentStreamCompressionByDefault();
        bookOptions.PageStartParity = null;
        ApplyLibraryDefaultFontEmbedding(options, bookOptions);
        using var book = PdfCore.PdfDocument.CreateForImportedLayout(bookOptions);
        book.Meta(title: title, author: author);
        if (!string.IsNullOrWhiteSpace(language)) book.Language(language!);
        var writerReport = new PdfCore.PdfConversionReport();
        if (options.PdfOptions.PageStartParity != null) {
            writerReport.Add(new PdfCore.PdfConversionWarning("OfficeIMO.Html.Pdf", "SectionStartParityUnavailable",
                "PageStartParity", "EPUB chapters start on the next physical page without parity padding.",
                PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission));
        }
        book.Options.ReportDiagnosticsTo(writerReport, "OfficeIMO.Html.Pdf");
        report(null, writerReport);
        var images = new PdfImageResourceCache(options);
        var pageReports = new List<PdfCore.PdfConversionReport>();
        return book.WritePreparedPages(destination, ReadPages(), (page, missing) => {
            pageReports[page].Add(new PdfCore.PdfConversionWarning("OfficeIMO.Html.Pdf", HtmlRenderDiagnosticCodes.HyperlinkTargetUnavailable,
                missing, "The target is outside the rendered content and its link annotation was omitted.",
                PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission));
        }, ResolveMissingSectionDestination, cancellationToken);

        IEnumerable<PdfCore.PageBlock> ReadPages() {
            int pageOffset = 0;
            int sectionOrder = 0;
            foreach (HtmlPdfSection section in sections) {
                cancellationToken.ThrowIfCancellationRequested();
                var sectionOptions = new HtmlPdfExportOptions(section.Request.Options) {
                    PageNumberOffset = pageOffset, IsPublicationSection = true
                };
                sectionOptions.PdfOptions.PageStartParity = null;
                PreparedPdfRequest prepared = PrepareRequest(section.Request.WithOptions(sectionOptions), sectionOptions,
                    allowNonPdfEncoder: true, preserveHtmlResourcePolicy: true);
                var effective = (HtmlPdfExportOptions)HtmlRenderEngine.PrepareOptions(section.Document, prepared.Request);
                HtmlRenderDocument layout = HtmlRenderEngine.ExecuteCore(section.Document, prepared.Request, effective, cancellationToken).Document;
                HtmlPdfRenderResult rendered = CreatePdf(layout, effective, cancellationToken, section.Outline, images, ++sectionOrder);
                using PdfCore.PdfDocument chapter = rendered.Document;
                PdfCore.PdfConversionReport sectionReport = HtmlPdfConverterExtensions.CollectRenderDiagnostics(rendered);
                report(section.Source, sectionReport);
                foreach (PdfCore.PageBlock page in chapter.Blocks.Cast<PdfCore.PageBlock>()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (pageOffset > 0) page.Options.ClearPageNumberStartOverride();
                    pageOffset++;
                    pageReports.Add(sectionReport);
                    yield return page;
                }
            }
        }
    }

    private static string? ResolveMissingSectionDestination(string name) {
        if (!name.StartsWith(NamedDestinationKeyPrefix, StringComparison.Ordinal)) return null;
        int separator = name.IndexOf(':', NamedDestinationKeyPrefix.Length);
        if (separator < 0 || !string.Equals(name.Substring(separator + 1), HtmlRenderOptions.PublicationFragmentPrefix + "top",
                StringComparison.OrdinalIgnoreCase)) return null;
        return name.Substring(0, separator + 1) + HtmlRenderOptions.PublicationStartAnchor;
    }

    private readonly struct HtmlPdfPageOutline {
        internal HtmlPdfPageOutline(HtmlPdfOutlineTarget target, int page, double y) {
            Target = target;
            Page = page;
            Y = y;
        }

        internal HtmlPdfOutlineTarget Target { get; }
        internal int Page { get; }
        internal double Y { get; }
    }

    private static ILookup<int, HtmlPdfPageOutline> ResolveDocumentOutlines(HtmlRenderDocument layout,
        IReadOnlyList<HtmlPdfOutlineTarget> outline, string prefix, HtmlDiagnosticReport diagnostics, CancellationToken token) {
        var destinations = new Dictionary<string, (int Page, double Y)>(StringComparer.Ordinal);
        if (outline.Any(target => !string.IsNullOrEmpty(target.Fragment))) {
            foreach (HtmlRenderPage page in layout.Pages) {
                foreach (HtmlRenderVisual visual in EnumerateVisuals(page.Scene)) {
                    token.ThrowIfCancellationRequested();
                    if (visual is HtmlRenderNamedDestination destination && !destinations.ContainsKey(destination.Name))
                        destinations.Add(destination.Name, (page.PageNumber, destination.Y));
                }
            }
        }
        return outline.Select(target => {
            token.ThrowIfCancellationRequested();
            if (string.IsNullOrEmpty(target.Fragment)) return new HtmlPdfPageOutline(target, 1, 0D);
            if (destinations.TryGetValue(prefix + HtmlRenderOptions.PublicationFragmentPrefix + target.Fragment, out var destination))
                return new HtmlPdfPageOutline(target, destination.Page, destination.Y);
            if (string.Equals(target.Fragment, "top", StringComparison.OrdinalIgnoreCase)) return new HtmlPdfPageOutline(target, 1, 0D);
            diagnostics.Add("OfficeIMO.Html.Pdf", "OutlineTargetUnavailable",
                "The outline target is outside the rendered content. The outline links to the start of the section.",
                HtmlDiagnosticSeverity.Warning, target.Fragment, lossKind: OfficeConversionLossKind.Approximation);
            return new HtmlPdfPageOutline(target, 1, 0D);
        }).ToLookup(item => item.Page);
    }
}
