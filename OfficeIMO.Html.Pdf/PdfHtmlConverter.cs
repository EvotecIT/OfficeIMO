using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// First-party PDF to HTML conversion helpers for the bidirectional HTML/PDF bridge.
/// </summary>
public static partial class PdfHtmlConverterExtensions {
    /// <summary>Renders an opened PDF as HTML.</summary>
    public static string ToHtml(this PdfCore.PdfDocument document, PdfHtmlSaveOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return ReadForHtml(document, options, options?.CancellationToken ?? default)
            .ToHtml(CreateRenderOptionsAfterPreselection(options));
    }

    /// <summary>Renders an opened PDF, saves the HTML as UTF-8 without a byte-order mark, and returns conversion diagnostics.</summary>
    public static PdfCore.PdfConversionReport SaveAsHtml(this PdfCore.PdfDocument document, string path, PdfHtmlSaveOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return ReadForHtml(document, options, options?.CancellationToken ?? default)
            .SaveAsHtml(path, CreateRenderOptionsAfterPreselection(options));
    }

    /// <summary>Renders an opened PDF, writes HTML to a caller-owned stream, and returns conversion diagnostics.</summary>
    public static PdfCore.PdfConversionReport SaveAsHtml(this PdfCore.PdfDocument document, Stream stream, PdfHtmlSaveOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return ReadForHtml(document, options, options?.CancellationToken ?? default)
            .SaveAsHtml(stream, CreateRenderOptionsAfterPreselection(options));
    }

    /// <summary>Renders an opened PDF, asynchronously saves the HTML, and returns conversion diagnostics.</summary>
    public static async Task<PdfCore.PdfConversionReport> SaveAsHtmlAsync(
        this PdfCore.PdfDocument document,
        string path,
        PdfHtmlSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        PdfHtmlSaveOptions renderOptions = CreateAsyncRenderOptions(options, cancellationToken, out CancellationTokenSource? linkedCancellation);
        using (linkedCancellation) {
            renderOptions.CancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfDocumentReadResult logical = ReadForHtml(document, renderOptions, renderOptions.CancellationToken);
            return await logical.SaveAsHtmlAsync(
                path,
                CreateRenderOptionsAfterPreselection(renderOptions),
                renderOptions.CancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>Renders an opened PDF, asynchronously writes HTML to a caller-owned stream, and returns conversion diagnostics.</summary>
    public static async Task<PdfCore.PdfConversionReport> SaveAsHtmlAsync(
        this PdfCore.PdfDocument document,
        Stream stream,
        PdfHtmlSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        PdfHtmlSaveOptions renderOptions = CreateAsyncRenderOptions(options, cancellationToken, out CancellationTokenSource? linkedCancellation);
        using (linkedCancellation) {
            renderOptions.CancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfDocumentReadResult logical = ReadForHtml(document, renderOptions, renderOptions.CancellationToken);
            return await logical.SaveAsHtmlAsync(
                stream,
                CreateRenderOptionsAfterPreselection(renderOptions),
                renderOptions.CancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Renders an already loaded logical PDF model as HTML.
    /// </summary>
    public static string ToHtml(this PdfCore.PdfDocumentReadResult document, PdfHtmlSaveOptions? options = null) {
        return document.ToHtmlResult(options).Value;
    }

    /// <summary>Renders a logical PDF document, saves the HTML as UTF-8 without a byte-order mark, and returns conversion diagnostics.</summary>
    public static PdfCore.PdfConversionReport SaveAsHtml(this PdfCore.PdfDocumentReadResult document, string path, PdfHtmlSaveOptions? options = null) {
        PdfHtmlConversionResult result = document.ToHtmlResult(options);
        HtmlTextIO.Write(path, result.Value);
        return result.Report;
    }

    /// <summary>Renders a logical PDF document, writes HTML to a caller-owned stream, and returns conversion diagnostics.</summary>
    public static PdfCore.PdfConversionReport SaveAsHtml(this PdfCore.PdfDocumentReadResult document, Stream stream, PdfHtmlSaveOptions? options = null) {
        PdfHtmlConversionResult result = document.ToHtmlResult(options);
        HtmlTextIO.Write(stream, result.Value);
        return result.Report;
    }

    /// <summary>Renders a logical PDF document, asynchronously saves the HTML, and returns conversion diagnostics.</summary>
    public static async Task<PdfCore.PdfConversionReport> SaveAsHtmlAsync(
        this PdfCore.PdfDocumentReadResult document,
        string path,
        PdfHtmlSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        PdfHtmlSaveOptions renderOptions = CreateAsyncRenderOptions(options, cancellationToken, out CancellationTokenSource? linkedCancellation);
        using (linkedCancellation) {
            renderOptions.CancellationToken.ThrowIfCancellationRequested();
            PdfHtmlConversionResult result = document.ToHtmlResult(renderOptions);
            await HtmlTextIO.WriteAsync(path, result.Value, renderOptions.CancellationToken).ConfigureAwait(false);
            return result.Report;
        }
    }

    /// <summary>Renders a logical PDF document, asynchronously writes HTML to a caller-owned stream, and returns conversion diagnostics.</summary>
    public static async Task<PdfCore.PdfConversionReport> SaveAsHtmlAsync(
        this PdfCore.PdfDocumentReadResult document,
        Stream stream,
        PdfHtmlSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        PdfHtmlSaveOptions renderOptions = CreateAsyncRenderOptions(options, cancellationToken, out CancellationTokenSource? linkedCancellation);
        using (linkedCancellation) {
            renderOptions.CancellationToken.ThrowIfCancellationRequested();
            PdfHtmlConversionResult result = document.ToHtmlResult(renderOptions);
            await HtmlTextIO.WriteAsync(stream, result.Value, renderOptions.CancellationToken).ConfigureAwait(false);
            return result.Report;
        }
    }

    internal static PdfHtmlSaveOptions CreateAsyncRenderOptions(
        PdfHtmlSaveOptions? options,
        CancellationToken cancellationToken,
        out CancellationTokenSource? linkedCancellation) {
        PdfHtmlSaveOptions renderOptions = options?.CloneForConversion() ?? new PdfHtmlSaveOptions();
        CancellationToken optionsCancellation = renderOptions.CancellationToken;
        linkedCancellation = null;

        if (optionsCancellation.CanBeCanceled && cancellationToken.CanBeCanceled && optionsCancellation != cancellationToken) {
            linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(optionsCancellation, cancellationToken);
            renderOptions.CancellationToken = linkedCancellation.Token;
        } else if (cancellationToken.CanBeCanceled) {
            renderOptions.CancellationToken = cancellationToken;
        }

        return renderOptions;
    }

    private static PdfCore.PdfPageRange[] CopyPageRanges(PdfHtmlSaveOptions options) {
        IReadOnlyList<PdfCore.PdfPageRange>? ranges = options.PageRanges;
        if (ranges == null || ranges.Count == 0) {
            return Array.Empty<PdfCore.PdfPageRange>();
        }

        var copy = new PdfCore.PdfPageRange[ranges.Count];
        for (int i = 0; i < ranges.Count; i++) {
            copy[i] = ranges[i];
        }

        return copy;
    }

    private static PdfCore.PdfDocumentReadResult ReadForHtml(
        PdfCore.PdfDocument document,
        PdfHtmlSaveOptions? options,
        CancellationToken cancellationToken = default) {
        PdfCore.PdfPageRange[] ranges = options is null
            ? Array.Empty<PdfCore.PdfPageRange>()
            : CopyPageRanges(options);
        PdfCore.PdfReadOptions configured = options?.ReadOptions ?? PdfCore.PdfReadOptions.Default;
        return document.Read(new PdfCore.PdfReadOptions {
            Profile = configured.Profile,
            PageSelection = ranges.Length == 0
                ? configured.PageSelection
                : PdfCore.PdfPageSelection.FromRanges(ranges),
            LayoutOptions = configured.LayoutOptions,
            Pipeline = configured.Pipeline
        }, cancellationToken);
    }

    private static PdfHtmlSaveOptions? CreateRenderOptionsAfterPreselection(PdfHtmlSaveOptions? options) {
        if (options?.PageRanges is null || options.PageRanges.Count == 0) return options;
        PdfHtmlSaveOptions renderOptions = options.CloneForConversion();
        renderOptions.PageRanges = null;
        return renderOptions;
    }

    private static string RenderSemanticDocument(PdfCore.PdfDocumentReadResult document, IReadOnlyList<PdfCore.PdfLogicalPage> pages, PdfHtmlSaveOptions options) {
        StringBuilder builder = CreateOutputBuilder(options);
        AppendDocumentStart(builder, document, options, positioned: false);
        if (options.EmitDocumentShell) {
            AppendBodyStart(builder, options, positioned: false);
        }

        if (options.IncludeMetadata) {
            AppendMetadataSection(builder, document);
        }

        AppendOutlineNavigation(builder, document, pages, options);
        AppendAcroFormXfaNotice(builder, document, options);

        for (int i = 0; i < pages.Count; i++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfLogicalPage page = pages[i];
            if (options.IncludePageContainers) {
                builder.Append("<section class=\"pdf-page\" id=\"");
                builder.Append(GetPageAnchorId(page.PageNumber, pages, i));
                builder.Append("\" data-page-number=\"");
                builder.Append(page.PageNumber.ToString(CultureInfo.InvariantCulture));
                builder.AppendLine("\">");
            }

            AppendSemanticPage(builder, page, options);

            if (options.IncludePageContainers) {
                builder.AppendLine("</section>");
            }
        }

        if (options.EmitDocumentShell) {
            builder.AppendLine("</body>");
            builder.AppendLine("</html>");
        }

        return NormalizeOutputNewLinesWithinBudget(builder, options);
    }

    private static string RenderPositionedReviewDocument(PdfCore.PdfDocumentReadResult document, IReadOnlyList<PdfCore.PdfLogicalPage> pages, PdfHtmlSaveOptions options) {
        StringBuilder builder = CreateOutputBuilder(options);
        AppendDocumentStart(builder, document, options, positioned: true);
        if (options.EmitDocumentShell) {
            AppendBodyStart(builder, options, positioned: true);
        } else {
            AppendPositionedStyles(builder, options.IncludeDefaultStyles);
        }

        AppendOutlineNavigation(builder, document, pages, options);
        AppendAcroFormXfaNotice(builder, document, options);

        for (int i = 0; i < pages.Count; i++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            AppendPositionedPage(builder, pages, i, options);
        }

        if (options.EmitDocumentShell) {
            builder.AppendLine("</body>");
            builder.AppendLine("</html>");
        }

        return NormalizeOutputNewLinesWithinBudget(builder, options);
    }

    private static IReadOnlyList<PdfCore.PdfLogicalPage> GetRenderPages(PdfCore.PdfDocumentReadResult document, PdfHtmlSaveOptions options) {
        PdfCore.PdfPageRange[] ranges = CopyPageRanges(options);
        if (ranges.Length == 0) {
            return document.Pages;
        }

        int maxSourcePageNumber = 0;
        for (int i = 0; i < document.Pages.Count; i++) {
            maxSourcePageNumber = Math.Max(maxSourcePageNumber, document.Pages[i].PageNumber);
        }

        if (maxSourcePageNumber == 0) {
            return Array.Empty<PdfCore.PdfLogicalPage>();
        }

        int[] pageNumbers = ExpandPageRanges(ranges, maxSourcePageNumber);
        var pages = new List<PdfCore.PdfLogicalPage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            IReadOnlyList<PdfCore.PdfLogicalPage> sourcePages = document.GetPages(pageNumbers[i]);
            for (int sourceIndex = 0; sourceIndex < sourcePages.Count; sourceIndex++) {
                pages.Add(sourcePages[sourceIndex]);
            }
        }

        return pages.AsReadOnly();
    }

    private static int[] ExpandPageRanges(PdfCore.PdfPageRange[] pageRanges, int pageCount) {
        if (pageRanges.Length == 0) {
            throw new ArgumentException("At least one page range must be specified.", nameof(PdfHtmlSaveOptions.PageRanges));
        }

        var pages = new List<int>();
        for (int i = 0; i < pageRanges.Length; i++) {
            PdfCore.PdfPageRange range = pageRanges[i];
            if (range.FirstPage < 1 || range.LastPage < range.FirstPage) {
                throw new ArgumentOutOfRangeException(nameof(PdfHtmlSaveOptions.PageRanges), "Page ranges must be inclusive one-based ranges.");
            }

            if (range.LastPage > pageCount) {
                throw new ArgumentOutOfRangeException(nameof(PdfHtmlSaveOptions.PageRanges), "Page range cannot exceed the document page count.");
            }

            for (int pageNumber = range.FirstPage; pageNumber <= range.LastPage; pageNumber++) {
                pages.Add(pageNumber);
            }
        }

        return pages.ToArray();
    }

    private static void AppendDocumentStart(StringBuilder builder, PdfCore.PdfDocumentReadResult document, PdfHtmlSaveOptions options, bool positioned) {
        if (!options.EmitDocumentShell) {
            return;
        }

        string title = string.IsNullOrWhiteSpace(document.Metadata.Title)
            ? options.DocumentTitleFallback
            : document.Metadata.Title!;
        builder.AppendLine("<!doctype html>");
        string? language = options.Language ?? document.CatalogLanguage;
        builder.Append("<html");
        if (!string.IsNullOrWhiteSpace(language)) {
            builder.Append(" lang=\"");
            builder.Append(HtmlAttribute(language!));
            builder.Append('"');
        }

        builder.AppendLine(">");
        builder.AppendLine("<head>");
        builder.AppendLine("<meta charset=\"utf-8\">");
        builder.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">");
        builder.Append("<title>");
        AppendHtmlText(builder, title);
        builder.AppendLine("</title>");
        if (options.IncludeMetadata) {
            AppendMeta(builder, "author", document.Metadata.Author);
            AppendMeta(builder, "description", document.Metadata.Subject);
            AppendMeta(builder, "keywords", document.Metadata.Keywords);
        }

        if (options.IncludeDefaultStyles) {
            builder.AppendLine("<style>");
            builder.AppendLine(OfficeHtmlDocumentShell.GetThemeCss(options.Theme));
            if (positioned) {
                builder.AppendLine(PdfHtmlReviewStyles.GetPositioning());
            }
            builder.AppendLine(PdfHtmlReviewStyles.GetReview());
            builder.AppendLine("</style>");
        } else if (positioned) {
            AppendPositionedStyles(builder, includeReviewStyles: false);
        }

        builder.AppendLine("</head>");
    }

    private static void AppendPositionedStyles(StringBuilder builder, bool includeReviewStyles = true) {
        builder.AppendLine("<style>");
        builder.AppendLine(PdfHtmlReviewStyles.GetPositioning());
        if (includeReviewStyles) {
            builder.AppendLine(PdfHtmlReviewStyles.GetReview());
        }
        builder.AppendLine("</style>");
    }

    private static void AppendBodyStart(StringBuilder builder, PdfHtmlSaveOptions options, bool positioned) {
        PdfHtmlProfileContract contract = PdfHtmlProfileContracts.Get(options.Profile);
        builder.Append("<body class=\"");
        builder.Append(HtmlAttribute(OfficeHtmlDocumentShell.MergeBodyClasses(
            "officeimo-html officeimo-pdf-html",
            positioned ? "officeimo-pdf-positioned" : "officeimo-pdf-semantic",
            options.DocumentOutput.BodyClass)));
        builder.Append("\" data-officeimo-html-profile=\"");
        builder.Append(HtmlAttribute(contract.Id));
        builder.Append("\" data-officeimo-html-theme=\"");
        builder.Append(HtmlAttribute(options.Theme.ToString()));
        builder.AppendLine("\">");
    }

    private static void AppendOutlineNavigation(StringBuilder builder, PdfCore.PdfDocumentReadResult document, IReadOnlyList<PdfCore.PdfLogicalPage> pages, PdfHtmlSaveOptions options) {
        if (!options.IncludeOutlines || document.Outlines.Count == 0) {
            return;
        }

        int renderedOutlineCount = CountRenderedOutlines(document, pages);
        if (renderedOutlineCount == 0) {
            return;
        }

        builder.Append("<nav class=\"pdf-outline\" aria-label=\"PDF outline\" data-outline-count=\"");
        builder.Append(CountOutlines(document.Outlines).ToString(CultureInfo.InvariantCulture));
        builder.Append("\" data-rendered-outline-count=\"");
        builder.Append(renderedOutlineCount.ToString(CultureInfo.InvariantCulture));
        builder.AppendLine("\">");
        builder.AppendLine("<ol>");
        AppendOutlineItems(builder, document.Outlines, document, pages);
        builder.AppendLine("</ol>");
        builder.AppendLine("</nav>");
    }

    private static void AppendOutlineItems(StringBuilder builder, IReadOnlyList<PdfCore.PdfOutlineItem> outlines, PdfCore.PdfDocumentReadResult document, IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
        for (int i = 0; i < outlines.Count; i++) {
            PdfCore.PdfOutlineItem outline = outlines[i];
            if (!ShouldRenderOutline(outline, document, pages)) {
                continue;
            }

            AppendOutlineItem(builder, outline, document, pages);
        }
    }

    private static void AppendOutlineItem(StringBuilder builder, PdfCore.PdfOutlineItem outline, PdfCore.PdfDocumentReadResult document, IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
        builder.Append("<li data-outline-level=\"");
        builder.Append(outline.Level.ToString(CultureInfo.InvariantCulture));
        builder.Append("\" data-expanded=\"");
        builder.Append(outline.IsExpanded ? "true" : "false");
        builder.Append('"');
        if (outline.PageNumber.HasValue) {
            builder.Append(" data-page-number=\"");
            builder.Append(outline.PageNumber.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append('"');
        }

        AppendOptionalDoubleAttribute(builder, "data-destination-top", outline.DestinationTop);
        AppendOptionalDoubleAttribute(builder, "data-destination-left", outline.DestinationLeft);
        AppendOptionalDoubleAttribute(builder, "data-destination-bottom", outline.DestinationBottom);
        AppendOptionalDoubleAttribute(builder, "data-destination-right", outline.DestinationRight);
        if (outline.DestinationMode.HasValue) {
            builder.Append(" data-destination-mode=\"");
            builder.Append(HtmlAttribute(outline.DestinationMode.Value.ToString()));
            builder.Append('"');
        }

        builder.Append('>');
        if (outline.PageNumber.HasValue && IsPageInRenderScope(outline.PageNumber.Value, pages)) {
            builder.Append("<a href=\"#");
            builder.Append(HtmlAttribute(GetFirstPageAnchorId(outline.PageNumber.Value, pages)));
            builder.Append("\">");
            AppendHtmlText(builder, outline.Title);
            builder.Append("</a>");
        } else {
            builder.Append("<span>");
            AppendHtmlText(builder, outline.Title);
            builder.Append("</span>");
        }

        if (HasRenderableOutlineChildren(outline, document, pages)) {
            builder.AppendLine();
            builder.AppendLine("<ol>");
            AppendOutlineItems(builder, outline.Children, document, pages);
            builder.AppendLine("</ol>");
        }

        builder.AppendLine("</li>");
    }

    private static void AppendOptionalDoubleAttribute(StringBuilder builder, string name, double? value) {
        if (!value.HasValue) {
            return;
        }

        builder.Append(' ');
        builder.Append(name);
        builder.Append("=\"");
        builder.Append(value.Value.ToString("0.###", CultureInfo.InvariantCulture));
        builder.Append('"');
    }

    private static int CountOutlines(IReadOnlyList<PdfCore.PdfOutlineItem> outlines) {
        int count = 0;
        for (int i = 0; i < outlines.Count; i++) {
            count++;
            count += CountOutlines(outlines[i].Children);
        }

        return count;
    }

    private static int CountRenderedOutlines(PdfCore.PdfDocumentReadResult document, IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
        int count = 0;
        CountRenderedOutlines(document.Outlines, document, pages, ref count);
        return count;
    }

    private static void CountRenderedOutlines(IReadOnlyList<PdfCore.PdfOutlineItem> outlines, PdfCore.PdfDocumentReadResult document, IReadOnlyList<PdfCore.PdfLogicalPage> pages, ref int count) {
        for (int i = 0; i < outlines.Count; i++) {
            PdfCore.PdfOutlineItem outline = outlines[i];
            if (!ShouldRenderOutline(outline, document, pages)) {
                continue;
            }

            count++;
            CountRenderedOutlines(outline.Children, document, pages, ref count);
        }
    }

    private static bool ShouldRenderOutline(PdfCore.PdfOutlineItem outline, PdfCore.PdfDocumentReadResult document, IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
        if (outline.PageNumber.HasValue) {
            return IsPageInRenderScope(outline.PageNumber.Value, pages) || HasRenderableOutlineChildren(outline, document, pages);
        }

        return AreAllDocumentPagesSelected(document, pages) || HasRenderableOutlineChildren(outline, document, pages);
    }

    private static bool HasRenderableOutlineChildren(PdfCore.PdfOutlineItem outline, PdfCore.PdfDocumentReadResult document, IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
        for (int i = 0; i < outline.Children.Count; i++) {
            if (ShouldRenderOutline(outline.Children[i], document, pages)) {
                return true;
            }
        }

        return false;
    }

    private static bool IsPageInRenderScope(int pageNumber, IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
        for (int i = 0; i < pages.Count; i++) {
            if (pages[i].PageNumber == pageNumber) {
                return true;
            }
        }

        return false;
    }

    private static string GetPageAnchorId(int pageNumber) =>
        "pdf-page-" + pageNumber.ToString(CultureInfo.InvariantCulture);

    private static string GetPageAnchorId(int pageNumber, IReadOnlyList<PdfCore.PdfLogicalPage> pages, int renderIndex) {
        int total = 0;
        int occurrence = 0;
        for (int i = 0; i < pages.Count; i++) {
            if (pages[i].PageNumber != pageNumber) {
                continue;
            }

            total++;
            if (i <= renderIndex) {
                occurrence++;
            }
        }

        if (total <= 1) {
            return GetPageAnchorId(pageNumber);
        }

        return GetPageAnchorId(pageNumber) + "-" + occurrence.ToString(CultureInfo.InvariantCulture);
    }

    private static string GetFirstPageAnchorId(int pageNumber, IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
        for (int i = 0; i < pages.Count; i++) {
            if (pages[i].PageNumber == pageNumber) {
                return GetPageAnchorId(pageNumber, pages, i);
            }
        }

        return GetPageAnchorId(pageNumber);
    }

    private static void AppendAcroFormXfaNotice(StringBuilder builder, PdfCore.PdfDocumentReadResult document, PdfHtmlSaveOptions options) {
        if (!document.HasAcroFormXfa || document.AcroFormXfa is null) {
            return;
        }

        AddWarning(
            options,
            "AcroFormXfaDetected",
            "AcroForm XFA packets are represented as inert review metadata; OfficeIMO.Html.Pdf does not render or fill XFA.",
            PdfCore.PdfConversionWarningSeverity.Warning);
        PdfCore.PdfAcroFormXfaInfo xfa = document.AcroFormXfa;
        builder.Append("<aside class=\"pdf-xfa-notice\" role=\"note\" data-xfa-object-kind=\"");
        builder.Append(HtmlAttribute(xfa.ObjectKind));
        builder.Append("\" data-xfa-packet-count=\"");
        builder.Append(xfa.PacketCount.ToString(CultureInfo.InvariantCulture));
        builder.Append("\" data-xfa-stream-count=\"");
        builder.Append(xfa.StreamCount.ToString(CultureInfo.InvariantCulture));
        builder.Append("\" data-xfa-payload-byte-count=\"");
        builder.Append(xfa.TotalPayloadBytes.ToString(CultureInfo.InvariantCulture));
        builder.Append('"');
        string? packetNames = FormatStringList(xfa.PacketNames, options);
        if (!string.IsNullOrWhiteSpace(packetNames)) {
            builder.Append(" data-xfa-packet-names=\"");
            builder.Append(HtmlAttribute(packetNames!));
            builder.Append('"');
        }

        builder.Append(">XFA form packets detected. OfficeIMO exposes packet metadata for review but does not render or fill XFA.</aside>");
        builder.AppendLine();
    }

    private static void AppendMeta(StringBuilder builder, string name, string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return;
        }

        builder.Append("<meta name=\"");
        builder.Append(HtmlAttribute(name));
        builder.Append("\" content=\"");
        builder.Append(HtmlAttribute(value!));
        builder.AppendLine("\">");
    }

    private static void AppendMetadataSection(StringBuilder builder, PdfCore.PdfDocumentReadResult document) {
        if (string.IsNullOrWhiteSpace(document.Metadata.Title) &&
            string.IsNullOrWhiteSpace(document.Metadata.Author) &&
            string.IsNullOrWhiteSpace(document.Metadata.Subject) &&
            string.IsNullOrWhiteSpace(document.Metadata.Keywords)) {
            return;
        }

        builder.AppendLine("<section class=\"pdf-metadata\">");
        if (!string.IsNullOrWhiteSpace(document.Metadata.Title)) {
            builder.Append("<h1>");
            AppendHtmlText(builder, document.Metadata.Title!);
            builder.AppendLine("</h1>");
        }

        AppendMetadataParagraph(builder, "Author", document.Metadata.Author);
        AppendMetadataParagraph(builder, "Subject", document.Metadata.Subject);
        AppendMetadataParagraph(builder, "Keywords", document.Metadata.Keywords);
        builder.AppendLine("</section>");
    }

    private static void AppendMetadataParagraph(StringBuilder builder, string label, string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return;
        }

        builder.Append("<p data-pdf-metadata=\"");
        builder.Append(HtmlAttribute(label));
        builder.Append("\"><strong>");
        AppendHtmlText(builder, label);
        builder.Append(":</strong> ");
        AppendHtmlText(builder, value!);
        builder.AppendLine("</p>");
    }

    private static string? FormatStringList(IReadOnlyList<string> values, PdfHtmlSaveOptions options) {
        if (values.Count == 0) {
            return null;
        }

        StringBuilder builder = CreateOutputBuilder(options);
        for (int i = 0; i < values.Count; i++) {
            if (string.IsNullOrWhiteSpace(values[i])) {
                continue;
            }

            if (builder.Length > 0) {
                builder.Append(',');
            }

            builder.Append(values[i]);
        }

        return builder.Length == 0 ? null : builder.ToString();
    }

    private static void AppendSemanticPage(StringBuilder builder, PdfCore.PdfLogicalPage page, PdfHtmlSaveOptions options) {
        List<HtmlItem> items = BuildSemanticPageItems(page, options);
        items.Sort(CompareHtmlItems);
        for (int i = 0; i < items.Count; i++) {
            builder.AppendLine(items[i].Html);
        }
    }

    private static List<HtmlItem> BuildSemanticPageItems(PdfCore.PdfLogicalPage page, PdfHtmlSaveOptions options) {
        var items = new List<HtmlItem>();
        IReadOnlyDictionary<(PdfCore.PdfLogicalReadingOrderKind Kind, int SourceIndex, int PlacementIndex), int> readingOrder =
            BuildReadingOrder(page, options.UseSharedPageReadingOrder);
        int sequence = 0;
        long retainedHtmlCharacters = 0L;

        for (int i = 0; i < page.Headings.Count; i++) {
            PdfCore.PdfLogicalHeading heading = page.Headings[i];
            int level = Math.Min(Math.Max(heading.Level, 1), 6);
            string html = RenderPageItemWithinBudget(options, retainedHtmlCharacters, builder => {
                builder.Append("<h").Append(level).Append('>');
                AppendHtmlText(builder, heading.Text);
                builder.Append("</h").Append(level).Append('>');
            });
            AddHtmlItem(items, new HtmlItem(heading.Line.BaselineY, heading.Line.XStart, sequence++, html, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.Heading, i)), options, ref retainedHtmlCharacters);
        }

        for (int i = 0; i < page.Paragraphs.Count; i++) {
            PdfCore.PdfLogicalParagraph paragraph = page.Paragraphs[i];
            if (IsParagraphRepresentedByStructuredElement(paragraph, page)) {
                continue;
            }

            string html = RenderPageItemWithinBudget(options, retainedHtmlCharacters, builder => {
                builder.Append("<p>");
                AppendHtmlText(builder, paragraph.Text);
                builder.Append("</p>");
            });
            AddHtmlItem(items, new HtmlItem(paragraph.YTop, paragraph.XStart, sequence++, html, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.Paragraph, i)), options, ref retainedHtmlCharacters);
        }

        for (int i = 0; i < page.ListItems.Count; i++) {
            PdfCore.PdfLogicalListItem listItem = page.ListItems[i];
            string html = RenderPageItemWithinBudget(options, retainedHtmlCharacters, builder => {
                builder.Append("<ul data-pdf-list-level=\"");
                builder.Append(Math.Max(1, listItem.Level).ToString(CultureInfo.InvariantCulture));
                builder.Append("\"><li>");
                AppendHtmlText(builder, listItem.Text);
                builder.Append("</li></ul>");
            });
            AddHtmlItem(items, new HtmlItem(listItem.Line.BaselineY, listItem.Line.XStart, sequence++, html, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.ListItem, i)), options, ref retainedHtmlCharacters);
        }

        for (int i = 0; i < page.Tables.Count; i++) {
            PdfCore.PdfLogicalTable table = page.Tables[i];
            string tableHtml = RenderSemanticTable(table, options, retainedHtmlCharacters);
            if (tableHtml.Length > 0) {
                double x = table.Columns.Count > 0 ? table.Columns[0].From : 0D;
                AddHtmlItem(items, new HtmlItem(table.YTop, x, sequence++, tableHtml, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.Table, i)), options, ref retainedHtmlCharacters);
            }
        }

        IReadOnlyList<PdfCore.IPdfLogicalElement> leaderRows = page.GetElements(PdfCore.PdfLogicalElementKind.LeaderRow);
        for (int i = 0; i < leaderRows.Count; i++) {
            if (leaderRows[i] is PdfCore.PdfLogicalLeaderRow leaderRow && !IsLeaderRowRepresentedByTable(leaderRow, page.Tables)) {
                string html = RenderPageItemWithinBudget(options, retainedHtmlCharacters, builder => {
                    builder.Append("<dl class=\"pdf-leader-row\"><dt>");
                    AppendHtmlText(builder, leaderRow.Label);
                    builder.Append("</dt><dd>");
                    AppendHtmlText(builder, leaderRow.Value);
                    builder.Append("</dd></dl>");
                });
                AddHtmlItem(items, new HtmlItem(null, 0D, sequence++, html), options, ref retainedHtmlCharacters);
            }
        }

        AppendUnmatchedTextBlocks(page, items, readingOrder, options, ref sequence, ref retainedHtmlCharacters);

        if (options.IncludeImagePlaceholders) {
            for (int i = 0; i < page.Images.Count; i++) {
                PdfCore.PdfLogicalImage image = page.Images[i];
                int placementIndex = image.Placements.Count == 0 ? -1 : 0;
                AddHtmlItem(items, new HtmlItem(null, 0D, sequence++, RenderImageFigure(image, options, retainedHtmlCharacters), GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.Image, i, placementIndex)), options, ref retainedHtmlCharacters);
            }
        }

        if (options.IncludeLinkAnnotations) {
            for (int i = 0; i < page.Links.Count; i++) {
                PdfCore.PdfLogicalLinkAnnotation link = page.Links[i];
                if (!HasHtmlLinkTarget(link)) {
                    continue;
                }

                string label = GetLinkLabel(link);
                string html = RenderPageItemWithinBudget(options, retainedHtmlCharacters, linkBuilder => {
                    if (link.Uri is not null && IsSafeLinkUri(link.Uri)) {
                        linkBuilder.Append("<p class=\"pdf-link\"><a");
                        AppendLinkTargetAttributes(linkBuilder, link);
                        linkBuilder.Append('>');
                        AppendHtmlText(linkBuilder, label);
                        linkBuilder.Append("</a></p>");
                    } else {
                        linkBuilder.Append("<p class=\"pdf-link\"");
                        AppendLinkTargetAttributes(linkBuilder, link);
                        linkBuilder.Append('>');
                        AppendHtmlText(linkBuilder, label);
                        linkBuilder.Append("</p>");
                    }
                });

                AddHtmlItem(items, new HtmlItem(link.Y2, link.X1, sequence++, html, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.Link, i)), options, ref retainedHtmlCharacters);
            }
        }

        if (options.IncludeFormWidgets) {
            for (int i = 0; i < page.FormWidgets.Count; i++) {
                PdfCore.PdfLogicalFormWidget widget = page.FormWidgets[i];
                string name = widget.FieldName ?? widget.FieldType ?? "Field";
                string value = widget.Value ?? string.Empty;
                string html = RenderPageItemWithinBudget(options, retainedHtmlCharacters, builder => {
                    builder.Append("<p class=\"pdf-form-widget\"><strong>");
                    AppendHtmlText(builder, name);
                    builder.Append("</strong>");
                    if (value.Length > 0) {
                        builder.Append(": ");
                        AppendHtmlText(builder, value);
                    }
                    builder.Append("</p>");
                });
                AddHtmlItem(items, new HtmlItem(widget.Y2, widget.X1, sequence++, html, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.FormWidget, i)), options, ref retainedHtmlCharacters);
            }
        }

        return items;
    }

    private static string RenderSemanticTable(
        PdfCore.PdfLogicalTable table,
        PdfHtmlSaveOptions options,
        long retainedHtmlCharacters) {
        if (table.Rows.Count == 0) {
            return string.Empty;
        }

        return RenderPageItemWithinBudget(options, retainedHtmlCharacters, builder => {
            builder.AppendLine("<table>");
            AppendTableRows(builder, table);
            builder.Append("</table>");
        });
    }

    private static void AppendTableRows(StringBuilder builder, PdfCore.PdfLogicalTable table) {
        PdfCore.PdfLogicalTableData data = PdfCore.PdfLogicalTableAnalysis.Extract(table);

        builder.Append("<tr>");
        for (int columnIndex = 0; columnIndex < data.Structure.ColumnCount; columnIndex++) {
            builder.Append("<th");
            if (data.IsNumericColumn(columnIndex)) {
                builder.Append(" class=\"pdf-numeric\" style=\"text-align:right\"");
            }

            builder.Append('>');
            AppendHtmlText(builder, columnIndex < data.Columns.Count ? data.Columns[columnIndex] : string.Empty);
            builder.Append("</th>");
        }

        builder.AppendLine("</tr>");

        for (int rowIndex = 0; rowIndex < data.Rows.Count; rowIndex++) {
            IReadOnlyList<string> row = data.Rows[rowIndex];
            builder.Append("<tr>");
            for (int columnIndex = 0; columnIndex < data.Structure.ColumnCount; columnIndex++) {
                builder.Append("<td");
                if (data.IsNumericColumn(columnIndex)) {
                    builder.Append(" class=\"pdf-numeric\" style=\"text-align:right\"");
                }

                builder.Append('>');
                AppendHtmlText(builder, columnIndex < row.Count ? row[columnIndex] : string.Empty);
                builder.Append("</td>");
            }

            builder.AppendLine("</tr>");
        }
    }

    private static string RenderImageFigure(
        PdfCore.PdfLogicalImage image,
        PdfHtmlSaveOptions options,
        long retainedHtmlCharacters) => RenderPageItemWithinBudget(options, retainedHtmlCharacters, builder => {
            builder.Append("<figure class=\"pdf-image-placeholder\" data-resource=\"");
            builder.Append(HtmlAttribute(image.ResourceName));
            builder.Append("\" data-page-number=\"");
            builder.Append(image.PageNumber.ToString(CultureInfo.InvariantCulture));
            builder.Append("\">");
            if (TryBuildEmbeddedImageDataUri(image, options, builder.MaxCapacity - builder.Length, out string? source)) {
                builder.Append("<img src=\"");
                builder.Append(HtmlAttribute(source!));
                builder.Append("\" alt=\"");
                builder.Append(HtmlAttribute("Image: " + image.ResourceName));
                builder.Append("\" width=\"");
                builder.Append(image.Width.ToString(CultureInfo.InvariantCulture));
                builder.Append("\" height=\"");
                builder.Append(image.Height.ToString(CultureInfo.InvariantCulture));
                builder.Append("\">");
            }

            builder.Append("<figcaption>Image: ");
            AppendHtmlText(builder, image.ResourceName);
            builder.Append(" (");
            builder.Append(image.Width.ToString(CultureInfo.InvariantCulture));
            builder.Append('x');
            builder.Append(image.Height.ToString(CultureInfo.InvariantCulture));
            if (!string.IsNullOrWhiteSpace(image.MimeType)) {
                builder.Append(", ");
                AppendHtmlText(builder, image.MimeType!);
            }

            builder.Append(")</figcaption></figure>");
        });

    private static bool TryBuildEmbeddedImageDataUri(
        PdfCore.PdfLogicalImage image,
        PdfHtmlSaveOptions options,
        int remainingItemCharacters,
        out string? source) {
        source = null;
        if (options.ImageExportMode != PdfHtmlImageExportMode.EmbeddedDataUri) {
            return false;
        }

        PdfCore.PdfExtractedImage sourceImage = image.SourceImage;
        if (!sourceImage.IsImageFile || string.IsNullOrWhiteSpace(sourceImage.MimeType)) {
            AddWarning(
                options,
                "ImageDataUnavailable",
                "An extracted PDF image was represented as a placeholder because it is not available as a complete image file.",
                PdfCore.PdfConversionWarningSeverity.Warning);
            return false;
        }

        if (options.MaxEmbeddedImageBytes.HasValue && sourceImage.Bytes.LongLength > options.MaxEmbeddedImageBytes.Value) {
            AddWarning(
                options,
                "ImageDataTooLarge",
                "An extracted PDF image was represented as a placeholder because it exceeds MaxEmbeddedImageBytes.",
                PdfCore.PdfConversionWarningSeverity.Warning);
            return false;
        }

        if (options.MaximumOutputCharacters.HasValue) {
            long base64Characters = ((sourceImage.Bytes.LongLength + 2L) / 3L) * 4L;
            long dataUriCharacters = "data:;base64,".Length + sourceImage.MimeType!.Length + base64Characters;
            if (dataUriCharacters > Math.Min(options.MaximumOutputCharacters.Value, remainingItemCharacters)) {
                AddWarning(
                    options,
                    "ImageDataTooLarge",
                    "An extracted PDF image was represented as a placeholder because its data URI exceeds MaximumOutputCharacters.",
                    PdfCore.PdfConversionWarningSeverity.Warning);
                return false;
            }
        }

        source = "data:" + sourceImage.MimeType + ";base64," + Convert.ToBase64String(sourceImage.Bytes);
        return true;
    }

    private static void AppendUnmatchedTextBlocks(
        PdfCore.PdfLogicalPage page,
        List<HtmlItem> items,
        IReadOnlyDictionary<(PdfCore.PdfLogicalReadingOrderKind Kind, int SourceIndex, int PlacementIndex), int> readingOrder,
        PdfHtmlSaveOptions options,
        ref int sequence,
        ref long retainedHtmlCharacters) {
        for (int i = 0; i < page.TextBlocks.Count; i++) {
            PdfCore.PdfLogicalTextBlock block = page.TextBlocks[i];
            if (IsTextBlockRepresented(block, page)) {
                continue;
            }

            AddHtmlItem(items, new HtmlItem(block.BaselineY, block.XStart, sequence++, RenderSemanticTextBlock(block, options, retainedHtmlCharacters), GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.TextBlock, i)), options, ref retainedHtmlCharacters);
        }
    }

    private static string RenderSemanticTextBlock(
        PdfCore.PdfLogicalTextBlock block,
        PdfHtmlSaveOptions options,
        long retainedHtmlCharacters) => RenderPageItemWithinBudget(options, retainedHtmlCharacters, builder => {
            (string Prefix, string Suffix) = block.Kind switch {
                PdfCore.PdfLogicalElementKind.Header => ("<header class=\"pdf-header\">", "</header>"),
                PdfCore.PdfLogicalElementKind.Footer => ("<footer class=\"pdf-footer\">", "</footer>"),
                PdfCore.PdfLogicalElementKind.Caption => ("<figure class=\"pdf-caption\"><figcaption>", "</figcaption></figure>"),
                PdfCore.PdfLogicalElementKind.Footnote => ("<aside class=\"pdf-footnote\" role=\"doc-footnote\">", "</aside>"),
                _ => ("<p>", "</p>")
            };
            builder.Append(Prefix);
            AppendHtmlText(builder, block.Text);
            builder.Append(Suffix);
        });

    private static bool IsTextBlockRepresented(PdfCore.PdfLogicalTextBlock block, PdfCore.PdfLogicalPage page) {
        if (block.Kind == PdfCore.PdfLogicalElementKind.Heading || block.Kind == PdfCore.PdfLogicalElementKind.ListItem) {
            return true;
        }

        for (int i = 0; i < page.Paragraphs.Count; i++) {
            PdfCore.PdfLogicalParagraph paragraph = page.Paragraphs[i];
            for (int lineIndex = 0; lineIndex < paragraph.Lines.Count; lineIndex++) {
                if (ReferenceEquals(paragraph.Lines[lineIndex], block)) {
                    return true;
                }
            }
        }

        for (int i = 0; i < page.Tables.Count; i++) {
            if (IsTextBlockRepresentedByTable(block, page.Tables[i])) {
                return true;
            }
        }

        return IsTextBlockRepresentedByLeaderRow(block, page);
    }

    private static bool IsParagraphRepresentedByStructuredElement(PdfCore.PdfLogicalParagraph paragraph, PdfCore.PdfLogicalPage page) {
        if (paragraph.Lines.Count == 0) {
            return false;
        }

        for (int i = 0; i < paragraph.Lines.Count; i++) {
            PdfCore.PdfLogicalTextBlock line = paragraph.Lines[i];
            bool represented = false;
            for (int tableIndex = 0; tableIndex < page.Tables.Count; tableIndex++) {
                if (IsTextBlockRepresentedByTable(line, page.Tables[tableIndex])) {
                    represented = true;
                    break;
                }
            }

            if (!represented && IsTextBlockRepresentedByLeaderRow(line, page)) {
                represented = true;
            }

            if (!represented) {
                return false;
            }
        }

        return true;
    }

    private static bool IsTextBlockRepresentedByTable(PdfCore.PdfLogicalTextBlock block, PdfCore.PdfLogicalTable table) {
        if (table.Rows.Count == 0 || table.Columns.Count == 0) {
            return false;
        }

        string blockText = NormalizeComparison(block.Text);
        if (blockText.Length == 0) {
            return false;
        }

        double blockLeft;
        double blockRight;
        if (block.VisualBounds is PdfCore.PdfLogicalVisualBounds blockBounds &&
            table.VisualBounds is PdfCore.PdfLogicalVisualBounds tableBounds) {
            double centerY = (blockBounds.Top + blockBounds.Bottom) / 2D;
            if (centerY < tableBounds.Top - 1D || centerY > tableBounds.Bottom + 1D) {
                return false;
            }

            blockLeft = Math.Min(blockBounds.Left, blockBounds.Right);
            blockRight = Math.Max(blockBounds.Left, blockBounds.Right);
        } else {
            if (table.VisualBounds is not null) return false;
            double top = Math.Max(table.YTop, table.YBottom);
            double bottom = Math.Min(table.YTop, table.YBottom);
            if (block.BaselineY > top + 1D || block.BaselineY < bottom - 1D) {
                return false;
            }

            blockLeft = Math.Min(block.XStart, block.XEnd);
            blockRight = Math.Max(block.XStart, block.XEnd);
        }

        var overlappingColumns = new List<int>();
        for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++) {
            PdfCore.PdfLogicalTableColumn column = table.Columns[columnIndex];
            double columnLeft = Math.Min(column.From, column.To);
            double columnRight = Math.Max(column.From, column.To);
            if (blockRight >= columnLeft - 1D && blockLeft <= columnRight + 1D) {
                overlappingColumns.Add(columnIndex);
            }
        }

        if (overlappingColumns.Count == 0) {
            return false;
        }

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            IReadOnlyList<string> row = table.Rows[rowIndex];
            var representedCells = new List<string>(overlappingColumns.Count);
            for (int columnIndex = 0; columnIndex < overlappingColumns.Count; columnIndex++) {
                int sourceColumnIndex = overlappingColumns[columnIndex];
                if (sourceColumnIndex < row.Count) {
                    representedCells.Add(row[sourceColumnIndex]);
                }
            }

            string representedText = NormalizeComparison(string.Join(" ", representedCells));
            if (representedText.Length == 0) {
                continue;
            }

            if (ContainsOrdinal(representedText, blockText) || ContainsOrdinal(blockText, representedText)) {
                return true;
            }
        }

        return false;
    }

    private static bool IsTextBlockRepresentedByLeaderRow(PdfCore.PdfLogicalTextBlock block, PdfCore.PdfLogicalPage page) {
        IReadOnlyList<PdfCore.IPdfLogicalElement> leaderRows = page.GetElements(PdfCore.PdfLogicalElementKind.LeaderRow);
        if (leaderRows.Count == 0) {
            return false;
        }

        string blockText = NormalizeComparison(block.Text);
        for (int i = 0; i < leaderRows.Count; i++) {
            if (leaderRows[i] is not PdfCore.PdfLogicalLeaderRow leaderRow) {
                continue;
            }

            string label = NormalizeComparison(leaderRow.Label);
            string value = NormalizeComparison(leaderRow.Value);
            if (label.Length > 0 && value.Length > 0 && ContainsOrdinal(blockText, label) && ContainsOrdinal(blockText, value)) {
                return true;
            }
        }

        return false;
    }

    private static bool IsLeaderRowRepresentedByTable(PdfCore.PdfLogicalLeaderRow leaderRow, IReadOnlyList<PdfCore.PdfLogicalTable> tables) {
        string label = NormalizeComparison(leaderRow.Label);
        string value = NormalizeComparison(leaderRow.Value);
        for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++) {
            PdfCore.PdfLogicalTable table = tables[tableIndex];
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                IReadOnlyList<string> row = table.Rows[rowIndex];
                if (row.Count >= 2 &&
                    NormalizeComparison(row[0]) == label &&
                    NormalizeComparison(row[row.Count - 1]) == value) {
                    return true;
                }
            }
        }

        return false;
    }

    private static int CompareHtmlItems(HtmlItem left, HtmlItem right) {
        if (left.ReadingOrderIndex.HasValue && right.ReadingOrderIndex.HasValue) {
            int orderComparison = left.ReadingOrderIndex.Value.CompareTo(right.ReadingOrderIndex.Value);
            if (orderComparison != 0) return orderComparison;
        } else if (left.ReadingOrderIndex.HasValue != right.ReadingOrderIndex.HasValue) {
            return left.ReadingOrderIndex.HasValue ? -1 : 1;
        }
        bool leftHasY = left.Y.HasValue;
        bool rightHasY = right.Y.HasValue;
        if (leftHasY && rightHasY) {
            int yComparison = right.Y!.Value.CompareTo(left.Y!.Value);
            if (yComparison != 0) {
                return yComparison;
            }

            int xComparison = left.X.CompareTo(right.X);
            if (xComparison != 0) {
                return xComparison;
            }
        } else if (leftHasY != rightHasY) {
            return leftHasY ? -1 : 1;
        }

        return left.Sequence.CompareTo(right.Sequence);
    }

    private static IReadOnlyDictionary<(PdfCore.PdfLogicalReadingOrderKind Kind, int SourceIndex, int PlacementIndex), int> BuildReadingOrder(
        PdfCore.PdfLogicalPage page,
        bool enabled) {
        if (!enabled) return new Dictionary<(PdfCore.PdfLogicalReadingOrderKind Kind, int SourceIndex, int PlacementIndex), int>();
        return PdfCore.PdfLogicalReadingOrderAnalysis.Analyze(
            page,
            PdfCore.PdfLogicalReadingOrderScope.PageContent).ToDictionary(
            static item => (item.Kind, item.SourceIndex, item.PlacementIndex),
            static item => item.OrderIndex);
    }

    private static int? GetReadingOrder(
        IReadOnlyDictionary<(PdfCore.PdfLogicalReadingOrderKind Kind, int SourceIndex, int PlacementIndex), int> readingOrder,
        PdfCore.PdfLogicalReadingOrderKind kind,
        int sourceIndex,
        int placementIndex = -1) => readingOrder.TryGetValue((kind, sourceIndex, placementIndex), out int index) ? index : null;

    private static string NormalizeComparison(string? text) {
        if (string.IsNullOrWhiteSpace(text)) {
            return string.Empty;
        }

        var builder = new StringBuilder(text!.Length);
        for (int i = 0; i < text.Length; i++) {
            char ch = text[i];
            if (!char.IsWhiteSpace(ch)) {
                builder.Append(char.ToUpperInvariant(ch));
            }
        }

        return builder.ToString();
    }

    private static bool ContainsOrdinal(string text, string value) {
        if (value.Length == 0) {
            return true;
        }

        if (value.Length > text.Length) {
            return false;
        }

        for (int i = 0; i <= text.Length - value.Length; i++) {
            if (string.Compare(text, i, value, 0, value.Length, StringComparison.Ordinal) == 0) {
                return true;
            }
        }

        return false;
    }

    private static string Points(double value) {
        return Math.Round(value, 3).ToString("0.###", CultureInfo.InvariantCulture) + "pt";
    }

    private static string FormatMatrix(PdfCore.PdfImagePlacement placement) {
        return string.Join(" ",
            placement.A.ToString("0.###", CultureInfo.InvariantCulture),
            placement.B.ToString("0.###", CultureInfo.InvariantCulture),
            placement.C.ToString("0.###", CultureInfo.InvariantCulture),
            placement.D.ToString("0.###", CultureInfo.InvariantCulture),
            placement.E.ToString("0.###", CultureInfo.InvariantCulture),
            placement.F.ToString("0.###", CultureInfo.InvariantCulture));
    }

    private static string HtmlAttribute(string value) {
        return System.Net.WebUtility.HtmlEncode(value ?? string.Empty).Replace("\"", "&quot;");
    }

    private static bool IsSafeLinkUri(string uri) {
        if (!Uri.TryCreate(uri, UriKind.Absolute, out Uri? parsed)) {
            return false;
        }

        return string.Equals(parsed.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(parsed.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(parsed.Scheme, Uri.UriSchemeMailto, StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasHtmlLinkTarget(PdfCore.PdfLogicalLinkAnnotation link) {
        return link.Uri is not null ||
            !string.IsNullOrWhiteSpace(link.DestinationName) ||
            link.DestinationPageNumber.HasValue;
    }

    private static string GetLinkLabel(PdfCore.PdfLogicalLinkAnnotation link) {
        if (!string.IsNullOrWhiteSpace(link.Contents)) {
            return link.Contents!;
        }

        if (!string.IsNullOrWhiteSpace(link.Uri)) {
            return link.Uri!;
        }

        if (!string.IsNullOrWhiteSpace(link.DestinationName)) {
            return link.DestinationName!;
        }

        if (link.DestinationPageNumber.HasValue) {
            return "Page " + link.DestinationPageNumber.Value.ToString(CultureInfo.InvariantCulture);
        }

        return "Link";
    }

    private static void AppendLinkTargetAttributes(StringBuilder builder, PdfCore.PdfLogicalLinkAnnotation link) {
        if (link.Uri is not null && IsSafeLinkUri(link.Uri)) {
            builder.Append(" href=\"");
            builder.Append(HtmlAttribute(link.Uri));
            builder.Append("\" rel=\"noopener noreferrer\"");
            return;
        }

        if (link.Uri is not null) {
            builder.Append(" data-unsafe-href=\"");
            builder.Append(HtmlAttribute(link.Uri));
            builder.Append('"');
            return;
        }

        if (!string.IsNullOrWhiteSpace(link.DestinationName)) {
            builder.Append(" data-destination=\"");
            builder.Append(HtmlAttribute(link.DestinationName!));
            builder.Append('"');
            return;
        }

        if (link.DestinationPageNumber.HasValue) {
            builder.Append(" data-destination-page-number=\"");
            builder.Append(link.DestinationPageNumber.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append('"');
            AppendOptionalDestinationAttribute(builder, "data-destination-mode", link.DestinationMode?.ToString());
            AppendOptionalDestinationAttribute(builder, "data-destination-left", link.DestinationLeft);
            AppendOptionalDestinationAttribute(builder, "data-destination-bottom", link.DestinationBottom);
            AppendOptionalDestinationAttribute(builder, "data-destination-right", link.DestinationRight);
            AppendOptionalDestinationAttribute(builder, "data-destination-top", link.DestinationTop);
        }
    }

    private static void AppendOptionalDestinationAttribute(StringBuilder builder, string name, string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return;
        }

        builder.Append(' ');
        builder.Append(name);
        builder.Append("=\"");
        builder.Append(HtmlAttribute(value!));
        builder.Append('"');
    }

    private static void AppendOptionalDestinationAttribute(StringBuilder builder, string name, double? value) {
        if (!value.HasValue) {
            return;
        }

        AppendOptionalDestinationAttribute(builder, name, value.Value.ToString("0.###", CultureInfo.InvariantCulture));
    }

    private static void AddWarning(
        PdfHtmlSaveOptions options,
        string code,
        string message,
        PdfCore.PdfConversionWarningSeverity severity) {
        options.Report.Add(new PdfCore.PdfConversionWarning(
            "OfficeIMO.Html.Pdf",
            code,
            "PDF to HTML",
            message,
            severity));
    }

    private sealed class HtmlItem {
        public HtmlItem(double? y, double x, int sequence, string html, int? readingOrderIndex = null) {
            Y = y;
            X = x;
            Sequence = sequence;
            Html = html;
            ReadingOrderIndex = readingOrderIndex;
        }

        public double? Y { get; }

        public double X { get; }

        public int Sequence { get; }

        public int? ReadingOrderIndex { get; }

        public string Html { get; }
    }
}
