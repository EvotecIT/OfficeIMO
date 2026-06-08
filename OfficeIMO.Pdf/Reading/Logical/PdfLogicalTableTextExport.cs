using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Options for exporting normalized logical PDF tables to lightweight text formats.
/// </summary>
public sealed class PdfLogicalTableTextExportOptions {
    /// <summary>
    /// PDF text layout options used when a path, stream, or byte array is loaded directly.
    /// </summary>
    public PdfTextLayoutOptions? LayoutOptions { get; set; }

    /// <summary>
    /// Optional inclusive one-based source page ranges used by direct PDF loading overloads.
    /// </summary>
    public IReadOnlyList<PdfPageRange>? PageRanges { get; set; }

    /// <summary>
    /// Maximum body rows to export per detected table. Values less than or equal to zero export all rows.
    /// </summary>
    public int MaxRows { get; set; }

    /// <summary>
    /// When true, each exported table includes a short source caption with PDF page and table number.
    /// </summary>
    public bool IncludeSourceCaptions { get; set; } = true;

    /// <summary>
    /// When true, Markdown exports insert the configured separator between tables from different source pages.
    /// </summary>
    public bool IncludePageSeparators { get; set; } = true;

    /// <summary>
    /// Markdown text used between pages when <see cref="IncludePageSeparators"/> is true.
    /// </summary>
    public string PageSeparator { get; set; } = "---";

    /// <summary>
    /// Right-align Markdown table columns when all non-empty body cells look numeric.
    /// </summary>
    public bool AlignNumericMarkdownColumns { get; set; } = true;

    /// <summary>
    /// When true, HTML exports include doctype, html, head, and body wrappers.
    /// </summary>
    public bool EmitHtmlDocumentShell { get; set; } = true;

    /// <summary>
    /// HTML document title used when <see cref="EmitHtmlDocumentShell"/> is true.
    /// </summary>
    public string HtmlDocumentTitle { get; set; } = "PDF Tables";

    /// <summary>
    /// Text emitted when no logical tables are found in the selected PDF pages.
    /// </summary>
    public string EmptyTableMessage { get; set; } = "No PDF tables detected.";

    /// <summary>
    /// When true, an empty-message document or fragment is emitted when no tables are found.
    /// </summary>
    public bool IncludeEmptyMessage { get; set; } = true;
}

/// <summary>
/// Table-focused Markdown and HTML export helpers over the first-party logical PDF read model.
/// </summary>
public static class PdfLogicalTableTextExportExtensions {
    /// <summary>
    /// Exports normalized logical PDF tables as Markdown tables.
    /// </summary>
    /// <param name="document">Logical PDF document to inspect.</param>
    /// <param name="options">Optional table export settings.</param>
    /// <returns>Markdown containing only detected PDF tables and optional source captions.</returns>
    public static string ToMarkdownTables(this PdfLogicalDocument document, PdfLogicalTableTextExportOptions? options = null) {
        Guard.NotNull(document, nameof(document));

        options ??= new PdfLogicalTableTextExportOptions();
        IReadOnlyList<PdfLogicalTableExtraction> tables = PdfLogicalTableAnalysis.ExtractTables(document, options.MaxRows);
        if (tables.Count == 0) {
            return options.IncludeEmptyMessage ? NormalizeEmptyMessage(options) : string.Empty;
        }

        var builder = new StringBuilder();
        int? previousPageNumber = null;
        for (int i = 0; i < tables.Count; i++) {
            PdfLogicalTableExtraction extraction = tables[i];
            if (builder.Length > 0) {
                builder.AppendLine();
                builder.AppendLine();
                if (options.IncludePageSeparators &&
                    previousPageNumber.HasValue &&
                    previousPageNumber.Value != extraction.PageNumber) {
                    builder.AppendLine(NormalizePageSeparator(options));
                    builder.AppendLine();
                }
            }

            if (options.IncludeSourceCaptions) {
                builder.Append("### ");
                builder.Append(BuildCaption(extraction));
                builder.AppendLine();
                builder.AppendLine();
            }

            AppendMarkdownTable(builder, extraction.Data, options);
            previousPageNumber = extraction.PageNumber;
        }

        return builder.ToString().TrimEnd();
    }

    /// <summary>
    /// Exports normalized logical PDF tables as semantic HTML tables.
    /// </summary>
    /// <param name="document">Logical PDF document to inspect.</param>
    /// <param name="options">Optional table export settings.</param>
    /// <returns>HTML containing only detected PDF tables and optional source captions.</returns>
    public static string ToHtmlTables(this PdfLogicalDocument document, PdfLogicalTableTextExportOptions? options = null) {
        Guard.NotNull(document, nameof(document));

        options ??= new PdfLogicalTableTextExportOptions();
        IReadOnlyList<PdfLogicalTableExtraction> tables = PdfLogicalTableAnalysis.ExtractTables(document, options.MaxRows);
        var builder = new StringBuilder();
        if (options.EmitHtmlDocumentShell) {
            AppendHtmlDocumentStart(builder, options);
        }

        if (tables.Count == 0) {
            if (options.IncludeEmptyMessage) {
                builder.Append("<p>");
                builder.Append(HtmlText(NormalizeEmptyMessage(options)));
                builder.AppendLine("</p>");
            }
        } else {
            for (int i = 0; i < tables.Count; i++) {
                AppendHtmlTable(builder, tables[i], options);
                if (i + 1 < tables.Count) {
                    builder.AppendLine();
                }
            }
        }

        if (options.EmitHtmlDocumentShell) {
            AppendHtmlDocumentEnd(builder);
        }

        return builder.ToString().TrimEnd();
    }

    /// <summary>
    /// Loads PDF bytes and exports detected logical tables as Markdown.
    /// </summary>
    /// <param name="pdfBytes">Source PDF bytes.</param>
    /// <param name="options">Optional table export settings.</param>
    /// <returns>Markdown containing only detected PDF tables.</returns>
    public static string ExtractMarkdownTables(byte[] pdfBytes, PdfLogicalTableTextExportOptions? options = null) {
        Guard.NotNull(pdfBytes, nameof(pdfBytes));

        options ??= new PdfLogicalTableTextExportOptions();
        return LoadPdf(pdfBytes, options).ToMarkdownTables(options);
    }

    /// <summary>
    /// Loads a PDF stream and exports detected logical tables as Markdown.
    /// </summary>
    /// <param name="pdfStream">Readable source PDF stream.</param>
    /// <param name="options">Optional table export settings.</param>
    /// <returns>Markdown containing only detected PDF tables.</returns>
    public static string ExtractMarkdownTables(Stream pdfStream, PdfLogicalTableTextExportOptions? options = null) {
        Guard.NotNull(pdfStream, nameof(pdfStream));

        options ??= new PdfLogicalTableTextExportOptions();
        return LoadPdf(pdfStream, options).ToMarkdownTables(options);
    }

    /// <summary>
    /// Loads a PDF file and exports detected logical tables as Markdown.
    /// </summary>
    /// <param name="pdfPath">Source PDF path.</param>
    /// <param name="options">Optional table export settings.</param>
    /// <returns>Markdown containing only detected PDF tables.</returns>
    public static string ExtractMarkdownTables(string pdfPath, PdfLogicalTableTextExportOptions? options = null) {
        if (string.IsNullOrWhiteSpace(pdfPath)) throw new ArgumentException("PDF path cannot be empty.", nameof(pdfPath));

        options ??= new PdfLogicalTableTextExportOptions();
        return LoadPdf(pdfPath, options).ToMarkdownTables(options);
    }

    /// <summary>
    /// Loads PDF bytes and exports detected logical tables as semantic HTML tables.
    /// </summary>
    /// <param name="pdfBytes">Source PDF bytes.</param>
    /// <param name="options">Optional table export settings.</param>
    /// <returns>HTML containing only detected PDF tables.</returns>
    public static string ExtractHtmlTables(byte[] pdfBytes, PdfLogicalTableTextExportOptions? options = null) {
        Guard.NotNull(pdfBytes, nameof(pdfBytes));

        options ??= new PdfLogicalTableTextExportOptions();
        return LoadPdf(pdfBytes, options).ToHtmlTables(options);
    }

    /// <summary>
    /// Loads a PDF stream and exports detected logical tables as semantic HTML tables.
    /// </summary>
    /// <param name="pdfStream">Readable source PDF stream.</param>
    /// <param name="options">Optional table export settings.</param>
    /// <returns>HTML containing only detected PDF tables.</returns>
    public static string ExtractHtmlTables(Stream pdfStream, PdfLogicalTableTextExportOptions? options = null) {
        Guard.NotNull(pdfStream, nameof(pdfStream));

        options ??= new PdfLogicalTableTextExportOptions();
        return LoadPdf(pdfStream, options).ToHtmlTables(options);
    }

    /// <summary>
    /// Loads a PDF file and exports detected logical tables as semantic HTML tables.
    /// </summary>
    /// <param name="pdfPath">Source PDF path.</param>
    /// <param name="options">Optional table export settings.</param>
    /// <returns>HTML containing only detected PDF tables.</returns>
    public static string ExtractHtmlTables(string pdfPath, PdfLogicalTableTextExportOptions? options = null) {
        if (string.IsNullOrWhiteSpace(pdfPath)) throw new ArgumentException("PDF path cannot be empty.", nameof(pdfPath));

        options ??= new PdfLogicalTableTextExportOptions();
        return LoadPdf(pdfPath, options).ToHtmlTables(options);
    }

    private static PdfLogicalDocument LoadPdf(string path, PdfLogicalTableTextExportOptions options) {
        PdfPageRange[] ranges = GetPageRanges(options);
        return ranges.Length == 0
            ? PdfLogicalDocument.Load(path, options.LayoutOptions)
            : PdfLogicalDocument.LoadPageRanges(path, options.LayoutOptions, ranges);
    }

    private static PdfLogicalDocument LoadPdf(byte[] pdfBytes, PdfLogicalTableTextExportOptions options) {
        PdfPageRange[] ranges = GetPageRanges(options);
        return ranges.Length == 0
            ? PdfLogicalDocument.Load(pdfBytes, options.LayoutOptions)
            : PdfLogicalDocument.LoadPageRanges(pdfBytes, options.LayoutOptions, ranges);
    }

    private static PdfLogicalDocument LoadPdf(Stream stream, PdfLogicalTableTextExportOptions options) {
        PdfPageRange[] ranges = GetPageRanges(options);
        return ranges.Length == 0
            ? PdfLogicalDocument.Load(stream, options.LayoutOptions)
            : PdfLogicalDocument.LoadPageRanges(stream, options.LayoutOptions, ranges);
    }

    private static PdfPageRange[] GetPageRanges(PdfLogicalTableTextExportOptions options) {
        return options.PageRanges == null || options.PageRanges.Count == 0
            ? Array.Empty<PdfPageRange>()
            : options.PageRanges.ToArray();
    }

    private static void AppendMarkdownTable(StringBuilder builder, PdfLogicalTableData data, PdfLogicalTableTextExportOptions options) {
        AppendMarkdownRow(builder, data.Columns, data.Structure.ColumnCount);
        builder.AppendLine();
        AppendMarkdownSeparator(builder, data, options.AlignNumericMarkdownColumns);

        for (int rowIndex = 0; rowIndex < data.Rows.Count; rowIndex++) {
            builder.AppendLine();
            AppendMarkdownRow(builder, data.Rows[rowIndex], data.Structure.ColumnCount);
        }
    }

    private static void AppendMarkdownRow(StringBuilder builder, IReadOnlyList<string> row, int columnCount) {
        builder.Append('|');
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            string cell = columnIndex < row.Count ? row[columnIndex] : string.Empty;
            builder.Append(' ');
            builder.Append(EscapeMarkdownTableCell(cell));
            builder.Append(" |");
        }
    }

    private static void AppendMarkdownSeparator(StringBuilder builder, PdfLogicalTableData data, bool alignNumericColumns) {
        builder.Append('|');
        for (int columnIndex = 0; columnIndex < data.Structure.ColumnCount; columnIndex++) {
            builder.Append(alignNumericColumns && data.IsNumericColumn(columnIndex) ? " ---: |" : " --- |");
        }
    }

    private static void AppendHtmlTable(StringBuilder builder, PdfLogicalTableExtraction extraction, PdfLogicalTableTextExportOptions options) {
        PdfLogicalTableData data = extraction.Data;
        builder.Append("<figure class=\"pdf-table\" data-page-number=\"");
        builder.Append(extraction.PageNumber.ToString(CultureInfo.InvariantCulture));
        builder.Append("\" data-table-index=\"");
        builder.Append(extraction.TableIndex.ToString(CultureInfo.InvariantCulture));
        builder.Append("\" data-detection-kind=\"");
        builder.Append(HtmlAttribute(extraction.DetectionKind));
        builder.Append('"');
        if (data.Truncated) {
            builder.Append(" data-truncated=\"true\" data-total-rows=\"");
            builder.Append(data.TotalRowCount.ToString(CultureInfo.InvariantCulture));
            builder.Append('"');
        }

        builder.AppendLine(">");

        if (options.IncludeSourceCaptions) {
            builder.Append("<figcaption>");
            builder.Append(HtmlText(BuildCaption(extraction)));
            builder.AppendLine("</figcaption>");
        }

        builder.AppendLine("<table>");
        AppendHtmlHeader(builder, data);
        AppendHtmlBody(builder, data);
        builder.AppendLine("</table>");
        builder.Append("</figure>");
    }

    private static void AppendHtmlHeader(StringBuilder builder, PdfLogicalTableData data) {
        builder.AppendLine("<thead>");
        builder.Append("<tr>");
        for (int columnIndex = 0; columnIndex < data.Structure.ColumnCount; columnIndex++) {
            AppendHtmlCellStart(builder, "th", data, columnIndex);
            builder.Append(HtmlText(columnIndex < data.Columns.Count ? data.Columns[columnIndex] : string.Empty));
            builder.Append("</th>");
        }

        builder.AppendLine("</tr>");
        builder.AppendLine("</thead>");
    }

    private static void AppendHtmlBody(StringBuilder builder, PdfLogicalTableData data) {
        builder.AppendLine("<tbody>");
        for (int rowIndex = 0; rowIndex < data.Rows.Count; rowIndex++) {
            IReadOnlyList<string> row = data.Rows[rowIndex];
            builder.Append("<tr>");
            for (int columnIndex = 0; columnIndex < data.Structure.ColumnCount; columnIndex++) {
                AppendHtmlCellStart(builder, "td", data, columnIndex);
                builder.Append(HtmlText(columnIndex < row.Count ? row[columnIndex] : string.Empty));
                builder.Append("</td>");
            }

            builder.AppendLine("</tr>");
        }

        builder.AppendLine("</tbody>");
    }

    private static void AppendHtmlCellStart(StringBuilder builder, string elementName, PdfLogicalTableData data, int columnIndex) {
        builder.Append('<');
        builder.Append(elementName);
        if (data.IsNumericColumn(columnIndex)) {
            builder.Append(" class=\"pdf-numeric\" style=\"text-align:right\"");
        }

        builder.Append('>');
    }

    private static void AppendHtmlDocumentStart(StringBuilder builder, PdfLogicalTableTextExportOptions options) {
        builder.AppendLine("<!doctype html>");
        builder.AppendLine("<html>");
        builder.AppendLine("<head>");
        builder.AppendLine("<meta charset=\"utf-8\">");
        builder.Append("<title>");
        builder.Append(HtmlText(string.IsNullOrWhiteSpace(options.HtmlDocumentTitle) ? "PDF Tables" : options.HtmlDocumentTitle));
        builder.AppendLine("</title>");
        builder.AppendLine("<style>table{border-collapse:collapse}th,td{border:1px solid #d0d7de;padding:4px 6px}.pdf-numeric{text-align:right}</style>");
        builder.AppendLine("</head>");
        builder.AppendLine("<body>");
    }

    private static void AppendHtmlDocumentEnd(StringBuilder builder) {
        builder.AppendLine();
        builder.AppendLine("</body>");
        builder.Append("</html>");
    }

    private static string BuildCaption(PdfLogicalTableExtraction extraction) {
        return "PDF page "
            + extraction.PageNumber.ToString(CultureInfo.InvariantCulture)
            + ", table "
            + (extraction.TableIndex + 1).ToString(CultureInfo.InvariantCulture);
    }

    private static string NormalizeEmptyMessage(PdfLogicalTableTextExportOptions options) {
        return string.IsNullOrWhiteSpace(options.EmptyTableMessage)
            ? "No PDF tables detected."
            : options.EmptyTableMessage.Trim();
    }

    private static string NormalizePageSeparator(PdfLogicalTableTextExportOptions options) {
        return string.IsNullOrWhiteSpace(options.PageSeparator)
            ? "---"
            : options.PageSeparator.Trim();
    }

    private static string EscapeMarkdownTableCell(string? text) {
        string value = NormalizeText(text);
        if (value.Length == 0) {
            return string.Empty;
        }

        var builder = new StringBuilder(value.Length + 8);
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (ch == '\\' ||
                ch == '`' ||
                ch == '*' ||
                ch == '_' ||
                ch == '[' ||
                ch == ']' ||
                ch == '<' ||
                ch == '>' ||
                ch == '|') {
                builder.Append('\\');
            }

            builder.Append(ch);
        }

        EscapeMarkdownLinePrefix(builder);
        return builder.ToString();
    }

    private static void EscapeMarkdownLinePrefix(StringBuilder builder) {
        int index = 0;
        while (index < builder.Length && char.IsWhiteSpace(builder[index])) {
            index++;
        }

        if (index >= builder.Length) {
            return;
        }

        char first = builder[index];
        if (first == '#' || first == '-' || first == '+' || first == '>') {
            builder.Insert(index, '\\');
            return;
        }

        if (!char.IsDigit(first)) {
            return;
        }

        int digitEnd = index + 1;
        while (digitEnd < builder.Length && char.IsDigit(builder[digitEnd])) {
            digitEnd++;
        }

        if (digitEnd < builder.Length && (builder[digitEnd] == '.' || builder[digitEnd] == ')')) {
            builder.Insert(digitEnd, '\\');
        }
    }

    private static string HtmlText(string? text) {
        return EscapeHtml(NormalizeText(text), escapeQuotes: false);
    }

    private static string HtmlAttribute(string? text) {
        return EscapeHtml(NormalizeText(text), escapeQuotes: true);
    }

    private static string EscapeHtml(string value, bool escapeQuotes) {
        if (value.Length == 0) {
            return string.Empty;
        }

        var builder = new StringBuilder(value.Length + 8);
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            switch (ch) {
                case '&':
                    builder.Append("&amp;");
                    break;
                case '<':
                    builder.Append("&lt;");
                    break;
                case '>':
                    builder.Append("&gt;");
                    break;
                case '"':
                    if (escapeQuotes) {
                        builder.Append("&quot;");
                    } else {
                        builder.Append('"');
                    }
                    break;
                case '\'':
                    if (escapeQuotes) {
                        builder.Append("&#39;");
                    } else {
                        builder.Append('\'');
                    }
                    break;
                default:
                    builder.Append(ch);
                    break;
            }
        }

        return builder.ToString();
    }

    private static string NormalizeText(string? text) {
        return string.IsNullOrWhiteSpace(text)
            ? string.Empty
            : text!.Replace("\r", " ").Replace("\n", " ").Trim();
    }
}
