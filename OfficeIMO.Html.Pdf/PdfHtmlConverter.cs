using System.Collections.Generic;
using System.Globalization;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// First-party PDF to HTML conversion helpers for the bidirectional HTML/PDF bridge.
/// </summary>
public static partial class PdfHtmlConverter {
    /// <summary>
    /// Converts PDF bytes to HTML using the first-party logical PDF read model.
    /// </summary>
    public static string ToHtml(byte[] pdf, PdfHtmlSaveOptions? options = null) {
        if (pdf == null) {
            throw new ArgumentNullException(nameof(pdf));
        }

        options ??= new PdfHtmlSaveOptions();
        options.ResetExportState();
        return RenderLogicalDocument(LoadLogical(pdf, options), options, applyPageRanges: false);
    }

    /// <summary>
    /// Converts a PDF file to HTML using the first-party logical PDF read model.
    /// </summary>
    public static string ToHtml(string path, PdfHtmlSaveOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new ArgumentException("PDF path cannot be empty.", nameof(path));
        }

        options ??= new PdfHtmlSaveOptions();
        options.ResetExportState();
        return RenderLogicalDocument(LoadLogical(path, options), options, applyPageRanges: false);
    }

    /// <summary>
    /// Converts PDF stream content to HTML using the first-party logical PDF read model.
    /// </summary>
    public static string ToHtml(Stream stream, PdfHtmlSaveOptions? options = null) {
        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        options ??= new PdfHtmlSaveOptions();
        options.ResetExportState();
        return RenderLogicalDocument(LoadLogical(stream, options), options, applyPageRanges: false);
    }

    /// <summary>
    /// Saves PDF bytes as an HTML file.
    /// </summary>
    public static void SaveAsHtml(byte[] pdf, string path, PdfHtmlSaveOptions? options = null) {
        File.WriteAllText(path, ToHtml(pdf, options), Encoding.UTF8);
    }

    /// <summary>
    /// Saves a PDF file as an HTML file.
    /// </summary>
    public static void SaveAsHtml(string pdfPath, string htmlPath, PdfHtmlSaveOptions? options = null) {
        File.WriteAllText(htmlPath, ToHtml(pdfPath, options), Encoding.UTF8);
    }

    /// <summary>
    /// Saves PDF stream content as an HTML file.
    /// </summary>
    public static void SaveAsHtml(Stream stream, string path, PdfHtmlSaveOptions? options = null) {
        File.WriteAllText(path, ToHtml(stream, options), Encoding.UTF8);
    }

    /// <summary>
    /// Renders an already parsed PDF document as HTML.
    /// </summary>
    public static string ToHtml(this PdfCore.PdfReadDocument document, PdfHtmlSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        options ??= new PdfHtmlSaveOptions();
        options.ResetExportState();
        return RenderLogicalDocument(LoadLogical(document, options), options, applyPageRanges: false);
    }

    /// <summary>
    /// Renders an already loaded logical PDF model as HTML.
    /// </summary>
    public static string ToHtml(this PdfCore.PdfLogicalDocument document, PdfHtmlSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        options ??= new PdfHtmlSaveOptions();
        options.ResetExportState();
        return RenderLogicalDocument(document, options, applyPageRanges: true);
    }

    private static string RenderLogicalDocument(PdfCore.PdfLogicalDocument document, PdfHtmlSaveOptions options, bool applyPageRanges) {
        IReadOnlyList<PdfCore.PdfLogicalPage> pages = applyPageRanges
            ? GetRenderPages(document, options)
            : document.Pages;
        return options.Profile switch {
            PdfHtmlProfile.Semantic => RenderSemanticDocument(document, pages, options),
            PdfHtmlProfile.PositionedReview => RenderPositionedReviewDocument(document, pages, options),
            _ => throw new ArgumentOutOfRangeException(nameof(options.Profile), options.Profile, "Unsupported PDF HTML profile.")
        };
    }

    private static PdfCore.PdfLogicalDocument LoadLogical(byte[] pdf, PdfHtmlSaveOptions options) {
        PdfCore.PdfPageRange[]? ranges = CopyPageRanges(options);
        PdfCore.PdfReadDocument document = PdfCore.PdfReadDocument.Load(pdf, options.ReadOptions);
        return ranges.Length > 0
            ? PdfCore.PdfLogicalDocument.FromPageRanges(document, options.LayoutOptions, ranges)
            : PdfCore.PdfLogicalDocument.From(document, options.LayoutOptions);
    }

    private static PdfCore.PdfLogicalDocument LoadLogical(string path, PdfHtmlSaveOptions options) {
        PdfCore.PdfPageRange[]? ranges = CopyPageRanges(options);
        PdfCore.PdfReadDocument document = PdfCore.PdfReadDocument.Load(path, options.ReadOptions);
        return ranges.Length > 0
            ? PdfCore.PdfLogicalDocument.FromPageRanges(document, options.LayoutOptions, ranges)
            : PdfCore.PdfLogicalDocument.From(document, options.LayoutOptions);
    }

    private static PdfCore.PdfLogicalDocument LoadLogical(Stream stream, PdfHtmlSaveOptions options) {
        PdfCore.PdfPageRange[]? ranges = CopyPageRanges(options);
        PdfCore.PdfReadDocument document = PdfCore.PdfReadDocument.Load(stream, options.ReadOptions);
        return ranges.Length > 0
            ? PdfCore.PdfLogicalDocument.FromPageRanges(document, options.LayoutOptions, ranges)
            : PdfCore.PdfLogicalDocument.From(document, options.LayoutOptions);
    }

    private static PdfCore.PdfLogicalDocument LoadLogical(PdfCore.PdfReadDocument document, PdfHtmlSaveOptions options) {
        PdfCore.PdfPageRange[]? ranges = CopyPageRanges(options);
        return ranges.Length > 0
            ? PdfCore.PdfLogicalDocument.FromPageRanges(document, options.LayoutOptions, ranges)
            : PdfCore.PdfLogicalDocument.From(document, options.LayoutOptions);
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

    private static string RenderSemanticDocument(PdfCore.PdfLogicalDocument document, IReadOnlyList<PdfCore.PdfLogicalPage> pages, PdfHtmlSaveOptions options) {
        var builder = new StringBuilder();
        AppendDocumentStart(builder, document, options, positioned: false);
        if (options.EmitDocumentShell) {
            builder.AppendLine("<body>");
        }

        if (options.IncludeMetadata) {
            AppendMetadataSection(builder, document);
        }

        AppendOutlineNavigation(builder, document, pages, options);
        AppendAcroFormXfaNotice(builder, document, options);

        for (int i = 0; i < pages.Count; i++) {
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

        return builder.ToString().TrimEnd();
    }

    private static string RenderPositionedReviewDocument(PdfCore.PdfLogicalDocument document, IReadOnlyList<PdfCore.PdfLogicalPage> pages, PdfHtmlSaveOptions options) {
        var builder = new StringBuilder();
        AppendDocumentStart(builder, document, options, positioned: true);
        if (options.EmitDocumentShell) {
            builder.AppendLine("<body>");
        } else {
            AppendPositionedStyles(builder);
        }

        AppendOutlineNavigation(builder, document, pages, options);
        AppendAcroFormXfaNotice(builder, document, options);

        for (int i = 0; i < pages.Count; i++) {
            AppendPositionedPage(builder, pages, i, options);
        }

        if (options.EmitDocumentShell) {
            builder.AppendLine("</body>");
            builder.AppendLine("</html>");
        }

        return builder.ToString().TrimEnd();
    }

    private static IReadOnlyList<PdfCore.PdfLogicalPage> GetRenderPages(PdfCore.PdfLogicalDocument document, PdfHtmlSaveOptions options) {
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

    private static void AppendDocumentStart(StringBuilder builder, PdfCore.PdfLogicalDocument document, PdfHtmlSaveOptions options, bool positioned) {
        if (!options.EmitDocumentShell) {
            return;
        }

        string title = string.IsNullOrWhiteSpace(document.Metadata.Title)
            ? options.DocumentTitleFallback
            : document.Metadata.Title!;
        builder.AppendLine("<!doctype html>");
        string? language = document.CatalogLanguage;
        builder.Append("<html");
        if (!string.IsNullOrWhiteSpace(language)) {
            builder.Append(" lang=\"");
            builder.Append(HtmlAttribute(language!));
            builder.Append('"');
        }

        builder.AppendLine(">");
        builder.AppendLine("<head>");
        builder.AppendLine("<meta charset=\"utf-8\">");
        builder.Append("<title>");
        builder.Append(HtmlText(title));
        builder.AppendLine("</title>");
        if (options.IncludeMetadata) {
            AppendMeta(builder, "author", document.Metadata.Author);
            AppendMeta(builder, "description", document.Metadata.Subject);
            AppendMeta(builder, "keywords", document.Metadata.Keywords);
        }

        if (positioned) {
            AppendPositionedStyles(builder);
        }

        builder.AppendLine("</head>");
    }

    private static void AppendPositionedStyles(StringBuilder builder) {
        builder.AppendLine("<style>");
        builder.AppendLine(".pdf-page{position:relative;margin:1rem auto;border:1px solid #cbd5e1;background:#fff;box-sizing:border-box;overflow:hidden;}");
        builder.AppendLine(".pdf-text{position:absolute;white-space:pre;line-height:1.2;font-family:Arial,sans-serif;font-size:10pt;}");
        builder.AppendLine(".pdf-heading{font-weight:700;}");
        builder.AppendLine(".pdf-list-item{padding-left:0.5rem;}");
        builder.AppendLine(".pdf-table{position:absolute;border-collapse:collapse;font-family:Arial,sans-serif;font-size:10pt;}");
        builder.AppendLine(".pdf-table td,.pdf-table th{border:1px solid #cbd5e1;padding:2pt 4pt;}");
        builder.AppendLine(".pdf-link,.pdf-form-widget{position:absolute;border:1px dashed #2563eb;background:rgba(37,99,235,.08);font-size:8pt;overflow:hidden;}");
        builder.AppendLine(".pdf-image-placeholder{font:8pt Arial,sans-serif;color:#475569;border:1px dashed #64748b;background:rgba(100,116,139,.08);box-sizing:border-box;overflow:hidden;}");
        builder.AppendLine(".pdf-outline{margin:1rem auto;max-width:60rem;font:9pt Arial,sans-serif;box-sizing:border-box;}");
        builder.AppendLine(".pdf-outline ol{margin:.25rem 0 .25rem 1.25rem;padding:0;}");
        builder.AppendLine(".pdf-outline a{color:#1d4ed8;text-decoration:none;}");
        builder.AppendLine(".pdf-xfa-notice{margin:1rem auto;padding:.5rem .75rem;max-width:60rem;border:1px solid #b45309;background:#fffbeb;color:#713f12;font:9pt Arial,sans-serif;box-sizing:border-box;}");
        builder.AppendLine("</style>");
    }

    private static void AppendOutlineNavigation(StringBuilder builder, PdfCore.PdfLogicalDocument document, IReadOnlyList<PdfCore.PdfLogicalPage> pages, PdfHtmlSaveOptions options) {
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

    private static void AppendOutlineItems(StringBuilder builder, IReadOnlyList<PdfCore.PdfOutlineItem> outlines, PdfCore.PdfLogicalDocument document, IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
        for (int i = 0; i < outlines.Count; i++) {
            PdfCore.PdfOutlineItem outline = outlines[i];
            if (!ShouldRenderOutline(outline, document, pages)) {
                continue;
            }

            AppendOutlineItem(builder, outline, document, pages);
        }
    }

    private static void AppendOutlineItem(StringBuilder builder, PdfCore.PdfOutlineItem outline, PdfCore.PdfLogicalDocument document, IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
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
            builder.Append(HtmlText(outline.Title));
            builder.Append("</a>");
        } else {
            builder.Append("<span>");
            builder.Append(HtmlText(outline.Title));
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

    private static int CountRenderedOutlines(PdfCore.PdfLogicalDocument document, IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
        int count = 0;
        CountRenderedOutlines(document.Outlines, document, pages, ref count);
        return count;
    }

    private static void CountRenderedOutlines(IReadOnlyList<PdfCore.PdfOutlineItem> outlines, PdfCore.PdfLogicalDocument document, IReadOnlyList<PdfCore.PdfLogicalPage> pages, ref int count) {
        for (int i = 0; i < outlines.Count; i++) {
            PdfCore.PdfOutlineItem outline = outlines[i];
            if (!ShouldRenderOutline(outline, document, pages)) {
                continue;
            }

            count++;
            CountRenderedOutlines(outline.Children, document, pages, ref count);
        }
    }

    private static bool ShouldRenderOutline(PdfCore.PdfOutlineItem outline, PdfCore.PdfLogicalDocument document, IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
        if (outline.PageNumber.HasValue) {
            return IsPageInRenderScope(outline.PageNumber.Value, pages) || HasRenderableOutlineChildren(outline, document, pages);
        }

        return AreAllDocumentPagesSelected(document, pages) || HasRenderableOutlineChildren(outline, document, pages);
    }

    private static bool HasRenderableOutlineChildren(PdfCore.PdfOutlineItem outline, PdfCore.PdfLogicalDocument document, IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
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

    private static void AppendAcroFormXfaNotice(StringBuilder builder, PdfCore.PdfLogicalDocument document, PdfHtmlSaveOptions options) {
        if (!document.HasAcroFormXfa || document.AcroFormXfa is null) {
            return;
        }

        AddWarning(options, "AcroFormXfaDetected", "AcroForm XFA packets are represented as inert review metadata; OfficeIMO.Html.Pdf does not render or fill XFA.");
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
        string? packetNames = FormatStringList(xfa.PacketNames);
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

    private static void AppendMetadataSection(StringBuilder builder, PdfCore.PdfLogicalDocument document) {
        if (string.IsNullOrWhiteSpace(document.Metadata.Title) &&
            string.IsNullOrWhiteSpace(document.Metadata.Author) &&
            string.IsNullOrWhiteSpace(document.Metadata.Subject) &&
            string.IsNullOrWhiteSpace(document.Metadata.Keywords)) {
            return;
        }

        builder.AppendLine("<section class=\"pdf-metadata\">");
        if (!string.IsNullOrWhiteSpace(document.Metadata.Title)) {
            builder.Append("<h1>");
            builder.Append(HtmlText(document.Metadata.Title!));
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
        builder.Append(HtmlText(label));
        builder.Append(":</strong> ");
        builder.Append(HtmlText(value!));
        builder.AppendLine("</p>");
    }

    private static string? FormatStringList(IReadOnlyList<string> values) {
        if (values.Count == 0) {
            return null;
        }

        var builder = new StringBuilder();
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
        int sequence = 0;

        for (int i = 0; i < page.Headings.Count; i++) {
            PdfCore.PdfLogicalHeading heading = page.Headings[i];
            int level = Math.Min(Math.Max(heading.Level, 1), 6);
            items.Add(new HtmlItem(heading.Line.BaselineY, heading.Line.XStart, sequence++, "<h" + level + ">" + HtmlText(heading.Text) + "</h" + level + ">"));
        }

        for (int i = 0; i < page.Paragraphs.Count; i++) {
            PdfCore.PdfLogicalParagraph paragraph = page.Paragraphs[i];
            if (IsParagraphRepresentedByStructuredElement(paragraph, page)) {
                continue;
            }

            items.Add(new HtmlItem(paragraph.YTop, paragraph.XStart, sequence++, "<p>" + HtmlText(paragraph.Text) + "</p>"));
        }

        for (int i = 0; i < page.ListItems.Count; i++) {
            PdfCore.PdfLogicalListItem listItem = page.ListItems[i];
            items.Add(new HtmlItem(listItem.Line.BaselineY, listItem.Line.XStart, sequence++, "<ul data-pdf-list-level=\"" + Math.Max(1, listItem.Level).ToString(CultureInfo.InvariantCulture) + "\"><li>" + HtmlText(listItem.Text) + "</li></ul>"));
        }

        for (int i = 0; i < page.Tables.Count; i++) {
            PdfCore.PdfLogicalTable table = page.Tables[i];
            string tableHtml = RenderSemanticTable(table);
            if (tableHtml.Length > 0) {
                double x = table.Columns.Count > 0 ? table.Columns[0].From : 0D;
                items.Add(new HtmlItem(table.YTop, x, sequence++, tableHtml));
            }
        }

        IReadOnlyList<PdfCore.IPdfLogicalElement> leaderRows = page.GetElements(PdfCore.PdfLogicalElementKind.LeaderRow);
        for (int i = 0; i < leaderRows.Count; i++) {
            if (leaderRows[i] is PdfCore.PdfLogicalLeaderRow leaderRow && !IsLeaderRowRepresentedByTable(leaderRow, page.Tables)) {
                items.Add(new HtmlItem(null, 0D, sequence++, "<dl class=\"pdf-leader-row\"><dt>" + HtmlText(leaderRow.Label) + "</dt><dd>" + HtmlText(leaderRow.Value) + "</dd></dl>"));
            }
        }

        AppendUnmatchedTextBlocks(page, items, ref sequence);

        if (options.IncludeImagePlaceholders) {
            for (int i = 0; i < page.Images.Count; i++) {
                PdfCore.PdfLogicalImage image = page.Images[i];
                items.Add(new HtmlItem(null, 0D, sequence++, RenderImageFigure(image, options)));
            }
        }

        if (options.IncludeLinkAnnotations) {
            for (int i = 0; i < page.Links.Count; i++) {
                PdfCore.PdfLogicalLinkAnnotation link = page.Links[i];
                if (!HasHtmlLinkTarget(link)) {
                    continue;
                }

                string label = GetLinkLabel(link);
                string html;
                if (link.Uri is not null && IsSafeLinkUri(link.Uri)) {
                    var linkBuilder = new StringBuilder();
                    linkBuilder.Append("<p class=\"pdf-link\"><a");
                    AppendLinkTargetAttributes(linkBuilder, link);
                    linkBuilder.Append('>');
                    linkBuilder.Append(HtmlText(label));
                    linkBuilder.Append("</a></p>");
                    html = linkBuilder.ToString();
                } else {
                    var linkBuilder = new StringBuilder();
                    linkBuilder.Append("<p class=\"pdf-link\"");
                    AppendLinkTargetAttributes(linkBuilder, link);
                    linkBuilder.Append('>');
                    linkBuilder.Append(HtmlText(label));
                    linkBuilder.Append("</p>");
                    html = linkBuilder.ToString();
                }

                items.Add(new HtmlItem(link.Y2, link.X1, sequence++, html));
            }
        }

        if (options.IncludeFormWidgets) {
            for (int i = 0; i < page.FormWidgets.Count; i++) {
                PdfCore.PdfLogicalFormWidget widget = page.FormWidgets[i];
                string name = widget.FieldName ?? widget.FieldType ?? "Field";
                string value = widget.Value ?? string.Empty;
                items.Add(new HtmlItem(widget.Y2, widget.X1, sequence++, "<p class=\"pdf-form-widget\"><strong>" + HtmlText(name) + "</strong>" + (value.Length > 0 ? ": " + HtmlText(value) : string.Empty) + "</p>"));
            }
        }

        return items;
    }

    private static string RenderSemanticTable(PdfCore.PdfLogicalTable table) {
        if (table.Rows.Count == 0) {
            return string.Empty;
        }

        var builder = new StringBuilder();
        builder.AppendLine("<table>");
        AppendTableRows(builder, table);
        builder.Append("</table>");
        return builder.ToString();
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
            builder.Append(HtmlText(columnIndex < data.Columns.Count ? data.Columns[columnIndex] : string.Empty));
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
                builder.Append(HtmlText(columnIndex < row.Count ? row[columnIndex] : string.Empty));
                builder.Append("</td>");
            }

            builder.AppendLine("</tr>");
        }
    }

    private static string RenderImageFigure(PdfCore.PdfLogicalImage image, PdfHtmlSaveOptions options) {
        var builder = new StringBuilder();
        builder.Append("<figure class=\"pdf-image-placeholder\" data-resource=\"");
        builder.Append(HtmlAttribute(image.ResourceName));
        builder.Append("\" data-page-number=\"");
        builder.Append(image.PageNumber.ToString(CultureInfo.InvariantCulture));
        builder.Append("\">");
        if (TryBuildEmbeddedImageDataUri(image, options, out string? source)) {
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
        builder.Append(HtmlText(image.ResourceName));
        builder.Append(" (");
        builder.Append(image.Width.ToString(CultureInfo.InvariantCulture));
        builder.Append('x');
        builder.Append(image.Height.ToString(CultureInfo.InvariantCulture));
        if (!string.IsNullOrWhiteSpace(image.MimeType)) {
            builder.Append(", ");
            builder.Append(HtmlText(image.MimeType!));
        }

        builder.Append(")</figcaption></figure>");
        return builder.ToString();
    }

    private static bool TryBuildEmbeddedImageDataUri(PdfCore.PdfLogicalImage image, PdfHtmlSaveOptions options, out string? source) {
        source = null;
        if (options.ImageExportMode != PdfHtmlImageExportMode.EmbeddedDataUri) {
            return false;
        }

        PdfCore.PdfExtractedImage sourceImage = image.SourceImage;
        if (!sourceImage.IsImageFile || string.IsNullOrWhiteSpace(sourceImage.MimeType)) {
            AddWarning(options, "ImageDataUnavailable", "An extracted PDF image was represented as a placeholder because it is not available as a complete image file.");
            return false;
        }

        if (options.MaxEmbeddedImageBytes.HasValue && sourceImage.Bytes.LongLength > options.MaxEmbeddedImageBytes.Value) {
            AddWarning(options, "ImageDataTooLarge", "An extracted PDF image was represented as a placeholder because it exceeds MaxEmbeddedImageBytes.");
            return false;
        }

        source = "data:" + sourceImage.MimeType + ";base64," + Convert.ToBase64String(sourceImage.Bytes);
        return true;
    }

    private static void AppendUnmatchedTextBlocks(PdfCore.PdfLogicalPage page, List<HtmlItem> items, ref int sequence) {
        for (int i = 0; i < page.TextBlocks.Count; i++) {
            PdfCore.PdfLogicalTextBlock block = page.TextBlocks[i];
            if (IsTextBlockRepresented(block, page)) {
                continue;
            }

            items.Add(new HtmlItem(block.BaselineY, block.XStart, sequence++, "<p>" + HtmlText(block.Text) + "</p>"));
        }
    }

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
        double top = Math.Max(table.YTop, table.YBottom);
        double bottom = Math.Min(table.YTop, table.YBottom);
        if (block.BaselineY > top + 1D || block.BaselineY < bottom - 1D) {
            return false;
        }

        string blockText = NormalizeComparison(block.Text);
        if (blockText.Length == 0) {
            return true;
        }

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            string rowText = NormalizeComparison(string.Join(" ", table.Rows[rowIndex]));
            if (rowText.Length == 0) {
                continue;
            }

            if (ContainsOrdinal(rowText, blockText) || ContainsOrdinal(blockText, rowText)) {
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

    private static string HtmlText(string value) {
        return System.Net.WebUtility.HtmlEncode(value ?? string.Empty);
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

    private static void AddWarning(PdfHtmlSaveOptions options, string code, string message) {
        options.ConversionReport.Add(new PdfCore.PdfConversionWarning(
            "OfficeIMO.Html.Pdf",
            code,
            "PDF to HTML",
            message,
            PdfCore.PdfConversionWarningSeverity.Information));
    }

    private sealed class HtmlItem {
        public HtmlItem(double? y, double x, int sequence, string html) {
            Y = y;
            X = x;
            Sequence = sequence;
            Html = html;
        }

        public double? Y { get; }

        public double X { get; }

        public int Sequence { get; }

        public string Html { get; }
    }
}
