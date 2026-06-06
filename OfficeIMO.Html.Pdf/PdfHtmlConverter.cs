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
        return LoadLogical(pdf, options).ToHtml(options);
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
        return LoadLogical(path, options).ToHtml(options);
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
        return LoadLogical(stream, options).ToHtml(options);
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
        return LoadLogical(document, options).ToHtml(options);
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
        return options.Profile switch {
            PdfHtmlProfile.Semantic => RenderSemanticDocument(document, options),
            PdfHtmlProfile.PositionedReview => RenderPositionedReviewDocument(document, options),
            _ => throw new ArgumentOutOfRangeException(nameof(options.Profile), options.Profile, "Unsupported PDF HTML profile.")
        };
    }

    private static PdfCore.PdfLogicalDocument LoadLogical(byte[] pdf, PdfHtmlSaveOptions options) {
        PdfCore.PdfPageRange[]? ranges = CopyPageRanges(options);
        return ranges.Length > 0
            ? PdfCore.PdfLogicalDocument.LoadPageRanges(pdf, options.LayoutOptions, ranges)
            : PdfCore.PdfLogicalDocument.Load(pdf, options.LayoutOptions);
    }

    private static PdfCore.PdfLogicalDocument LoadLogical(string path, PdfHtmlSaveOptions options) {
        PdfCore.PdfPageRange[]? ranges = CopyPageRanges(options);
        return ranges.Length > 0
            ? PdfCore.PdfLogicalDocument.LoadPageRanges(path, options.LayoutOptions, ranges)
            : PdfCore.PdfLogicalDocument.Load(path, options.LayoutOptions);
    }

    private static PdfCore.PdfLogicalDocument LoadLogical(Stream stream, PdfHtmlSaveOptions options) {
        PdfCore.PdfPageRange[]? ranges = CopyPageRanges(options);
        return ranges.Length > 0
            ? PdfCore.PdfLogicalDocument.LoadPageRanges(stream, options.LayoutOptions, ranges)
            : PdfCore.PdfLogicalDocument.Load(stream, options.LayoutOptions);
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

    private static string RenderSemanticDocument(PdfCore.PdfLogicalDocument document, PdfHtmlSaveOptions options) {
        var builder = new StringBuilder();
        AppendDocumentStart(builder, document, options, positioned: false);
        if (options.EmitDocumentShell) {
            builder.AppendLine("<body>");
        }

        if (options.IncludeMetadata) {
            AppendMetadataSection(builder, document);
        }

        for (int i = 0; i < document.Pages.Count; i++) {
            PdfCore.PdfLogicalPage page = document.Pages[i];
            if (options.IncludePageContainers) {
                builder.Append("<section class=\"pdf-page\" data-page-number=\"");
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

    private static string RenderPositionedReviewDocument(PdfCore.PdfLogicalDocument document, PdfHtmlSaveOptions options) {
        var builder = new StringBuilder();
        AppendDocumentStart(builder, document, options, positioned: true);
        if (options.EmitDocumentShell) {
            builder.AppendLine("<body>");
        } else {
            AppendPositionedStyles(builder);
        }

        for (int i = 0; i < document.Pages.Count; i++) {
            AppendPositionedPage(builder, document.Pages[i], options);
        }

        if (options.EmitDocumentShell) {
            builder.AppendLine("</body>");
            builder.AppendLine("</html>");
        }

        return builder.ToString().TrimEnd();
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
        builder.AppendLine("</style>");
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
                string target = link.Uri ?? link.DestinationName ?? string.Empty;
                if (target.Length == 0) {
                    continue;
                }

                string label = !string.IsNullOrWhiteSpace(link.Contents) ? link.Contents! : target;
                string html;
                if (link.Uri is not null) {
                    html = IsSafeLinkUri(link.Uri)
                        ? "<p class=\"pdf-link\"><a href=\"" + HtmlAttribute(link.Uri) + "\">" + HtmlText(label) + "</a></p>"
                        : "<p class=\"pdf-link\" data-unsafe-href=\"" + HtmlAttribute(link.Uri) + "\">" + HtmlText(label) + "</p>";
                } else {
                    html = "<p class=\"pdf-link\" data-destination=\"" + HtmlAttribute(target) + "\">" + HtmlText(label) + "</p>";
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
        int columnCount = 0;
        for (int i = 0; i < table.Rows.Count; i++) {
            columnCount = Math.Max(columnCount, table.Rows[i].Count);
        }

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            IReadOnlyList<string> row = table.Rows[rowIndex];
            builder.Append("<tr>");
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                string tag = rowIndex == 0 ? "th" : "td";
                builder.Append('<');
                builder.Append(tag);
                builder.Append('>');
                builder.Append(HtmlText(columnIndex < row.Count ? row[columnIndex] : string.Empty));
                builder.Append("</");
                builder.Append(tag);
                builder.Append('>');
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
