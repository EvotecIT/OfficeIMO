using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Options for rendering a logical PDF read model as Markdown.
/// </summary>
public sealed class PdfLogicalMarkdownOptions {
    /// <summary>Emit horizontal-rule separators between logical pages.</summary>
    public bool IncludePageSeparators { get; set; } = true;

    /// <summary>Emit readable placeholders for image elements discovered in the logical model.</summary>
    public bool IncludeImagePlaceholders { get; set; } = true;

    /// <summary>Emit a link annotation section for supported URI, named-destination, direct-destination, named-action, and remote GoTo links.</summary>
    public bool IncludeLinkAnnotations { get; set; }

    /// <summary>Emit a form widget section for AcroForm widgets placed on pages.</summary>
    public bool IncludeFormWidgets { get; set; }

    /// <summary>Right-align table columns in Markdown when all non-empty body cells look numeric.</summary>
    public bool AlignNumericTableColumns { get; set; } = true;

    /// <summary>Markdown text used between pages when <see cref="IncludePageSeparators"/> is true.</summary>
    public string PageSeparator { get; set; } = "---";
}

/// <summary>
/// Markdown rendering helpers for the first-party logical PDF read model.
/// </summary>
public static class PdfLogicalMarkdownExtensions {
    /// <summary>
    /// Renders the logical PDF document as Markdown using headings, paragraphs, lists, tables, and optional annotations from the existing logical model.
    /// </summary>
    public static string ToMarkdown(this PdfLogicalDocument document, PdfLogicalMarkdownOptions? options = null) {
        Guard.NotNull(document, nameof(document));
        options ??= new PdfLogicalMarkdownOptions();

        var builder = new StringBuilder();
        for (int i = 0; i < document.Pages.Count; i++) {
            if (i > 0 && options.IncludePageSeparators) {
                AppendBlock(builder, options.PageSeparator);
            }

            AppendPage(builder, document.Pages[i], options);
        }

        return builder.ToString().TrimEnd();
    }

    /// <summary>
    /// Renders a logical PDF page as Markdown using headings, paragraphs, lists, tables, and optional annotations from the existing logical model.
    /// </summary>
    public static string ToMarkdown(this PdfLogicalPage page, PdfLogicalMarkdownOptions? options = null) {
        Guard.NotNull(page, nameof(page));
        options ??= new PdfLogicalMarkdownOptions();

        var builder = new StringBuilder();
        AppendPage(builder, page, options);
        return builder.ToString().TrimEnd();
    }

    private static void AppendPage(StringBuilder builder, PdfLogicalPage page, PdfLogicalMarkdownOptions options) {
        List<MarkdownItem> items = BuildPageItems(page, options);
        items.Sort(CompareMarkdownItems);

        for (int i = 0; i < items.Count; i++) {
            AppendBlock(builder, items[i].Markdown);
        }
    }

    private static List<MarkdownItem> BuildPageItems(PdfLogicalPage page, PdfLogicalMarkdownOptions options) {
        var items = new List<MarkdownItem>();
        int sequence = 0;

        for (int i = 0; i < page.Headings.Count; i++) {
            PdfLogicalHeading heading = page.Headings[i];
            int level = Math.Min(Math.Max(heading.Level, 1), 6);
            items.Add(new MarkdownItem(heading.Line.BaselineY, heading.Line.XStart, sequence++, new string('#', level) + " " + EscapeInline(heading.Text)));
        }

        for (int i = 0; i < page.Paragraphs.Count; i++) {
            PdfLogicalParagraph paragraph = page.Paragraphs[i];
            if (IsParagraphRepresentedByStructuredElement(paragraph, page)) {
                continue;
            }

            items.Add(new MarkdownItem(paragraph.YTop, paragraph.XStart, sequence++, EscapeInline(paragraph.Text)));
        }

        for (int i = 0; i < page.ListItems.Count; i++) {
            PdfLogicalListItem listItem = page.ListItems[i];
            string indent = new string(' ', Math.Max(listItem.Level - 1, 0) * 2);
            items.Add(new MarkdownItem(listItem.Line.BaselineY, listItem.Line.XStart, sequence++, indent + FormatListMarker(listItem.Marker) + " " + EscapeInline(listItem.Text)));
        }

        for (int i = 0; i < page.Tables.Count; i++) {
            PdfLogicalTable table = page.Tables[i];
            double x = table.Columns.Count > 0 ? table.Columns[0].From : 0;
            string markdown = RenderTable(table, options);
            if (markdown.Length > 0) {
                items.Add(new MarkdownItem(table.YTop, x, sequence++, markdown));
            }
        }

        IReadOnlyList<IPdfLogicalElement> leaderRows = page.GetElements(PdfLogicalElementKind.LeaderRow);
        for (int i = 0; i < leaderRows.Count; i++) {
            if (leaderRows[i] is PdfLogicalLeaderRow leaderRow) {
                if (IsLeaderRowRepresentedByTable(leaderRow, page.Tables)) {
                    continue;
                }

                items.Add(new MarkdownItem(null, 0, sequence++, EscapeInline(leaderRow.Label) + " | " + EscapeInline(leaderRow.Value)));
            }
        }

        AppendUnmatchedTextBlocks(page, items, ref sequence);

        if (options.IncludeImagePlaceholders) {
            for (int i = 0; i < page.Images.Count; i++) {
                PdfLogicalImage image = page.Images[i];
                string description = "[Image: page " + image.PageNumber + ", resource " + EscapeInline(image.ResourceName) + ", " + image.Width + "x" + image.Height;
                string? mimeType = image.MimeType;
                if (!string.IsNullOrEmpty(mimeType)) {
                    description += ", " + EscapeInline(mimeType!);
                }

                description += "]";
                items.Add(new MarkdownItem(null, 0, sequence++, description));
            }
        }

        if (options.IncludeLinkAnnotations) {
            for (int i = 0; i < page.Links.Count; i++) {
                PdfLogicalLinkAnnotation link = page.Links[i];
                string? target = FormatLinkAnnotationTarget(link);
                if (string.IsNullOrEmpty(target)) {
                    continue;
                }

                string label = !string.IsNullOrWhiteSpace(link.Contents) ? link.Contents! : target!;
                string markdown = link.Uri is not null
                    ? IsSafeLinkUri(link.Uri)
                        ? "[Link: " + EscapeInline(label) + "](" + EscapeLinkTarget(link.Uri) + ")"
                        : "[Link: " + EscapeInline(label) + " -> " + EscapeInline(link.Uri) + "]"
                    : "[Link: " + EscapeInline(label) + " -> " + EscapeInline(target!) + "]";
                items.Add(new MarkdownItem(link.Y2, link.X1, sequence++, markdown));
            }
        }

        if (options.IncludeFormWidgets) {
            for (int i = 0; i < page.FormWidgets.Count; i++) {
                PdfLogicalFormWidget widget = page.FormWidgets[i];
                string name = widget.FieldName ?? widget.FieldType ?? "Field";
                string value = widget.Value ?? string.Empty;
                items.Add(new MarkdownItem(widget.Y2, widget.X1, sequence++, "[Form field: " + EscapeInline(name) + (value.Length > 0 ? " = " + EscapeInline(value) : string.Empty) + "]"));
            }
        }

        return items;
    }

    private static void AppendUnmatchedTextBlocks(PdfLogicalPage page, List<MarkdownItem> items, ref int sequence) {
        for (int i = 0; i < page.TextBlocks.Count; i++) {
            PdfLogicalTextBlock block = page.TextBlocks[i];
            if (IsTextBlockRepresented(block, page)) {
                continue;
            }

            items.Add(new MarkdownItem(block.BaselineY, block.XStart, sequence++, EscapeInline(block.Text)));
        }
    }

    private static bool IsTextBlockRepresented(PdfLogicalTextBlock block, PdfLogicalPage page) {
        if (block.Kind == PdfLogicalElementKind.Heading || block.Kind == PdfLogicalElementKind.ListItem) {
            return true;
        }

        for (int i = 0; i < page.Paragraphs.Count; i++) {
            PdfLogicalParagraph paragraph = page.Paragraphs[i];
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

        if (IsTextBlockRepresentedByLeaderRow(block, page)) {
            return true;
        }

        return false;
    }

    private static bool IsParagraphRepresentedByStructuredElement(PdfLogicalParagraph paragraph, PdfLogicalPage page) {
        if (paragraph.Lines.Count == 0) {
            return false;
        }

        for (int i = 0; i < paragraph.Lines.Count; i++) {
            PdfLogicalTextBlock line = paragraph.Lines[i];
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

    private static bool IsTextBlockRepresentedByTable(PdfLogicalTextBlock block, PdfLogicalTable table) {
        double top = Math.Max(table.YTop, table.YBottom);
        double bottom = Math.Min(table.YTop, table.YBottom);
        if (block.BaselineY > top + 1D || block.BaselineY < bottom - 1D) {
            return false;
        }

        string blockText = NormalizeMarkdownComparison(block.Text);
        if (blockText.Length == 0) {
            return true;
        }

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            string rowText = NormalizeMarkdownComparison(string.Join(" ", table.Rows[rowIndex]));
            if (rowText.Length == 0) {
                continue;
            }

            if (ContainsOrdinal(rowText, blockText) ||
                ContainsOrdinal(blockText, rowText)) {
                return true;
            }
        }

        return false;
    }

    private static bool IsTextBlockRepresentedByLeaderRow(PdfLogicalTextBlock block, PdfLogicalPage page) {
        IReadOnlyList<IPdfLogicalElement> leaderRows = page.GetElements(PdfLogicalElementKind.LeaderRow);
        if (leaderRows.Count == 0) {
            return false;
        }

        string blockText = NormalizeMarkdownComparison(block.Text);
        for (int i = 0; i < leaderRows.Count; i++) {
            if (leaderRows[i] is not PdfLogicalLeaderRow leaderRow) {
                continue;
            }

            string label = NormalizeMarkdownComparison(leaderRow.Label);
            string value = NormalizeMarkdownComparison(leaderRow.Value);
            if (label.Length == 0 || value.Length == 0) {
                continue;
            }

            if (ContainsOrdinal(blockText, label) && ContainsOrdinal(blockText, value)) {
                return true;
            }
        }

        return false;
    }

    private static bool IsLeaderRowRepresentedByTable(PdfLogicalLeaderRow leaderRow, IReadOnlyList<PdfLogicalTable> tables) {
        string label = NormalizeMarkdownComparison(leaderRow.Label);
        string value = NormalizeMarkdownComparison(leaderRow.Value);
        for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++) {
            PdfLogicalTable table = tables[tableIndex];
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                IReadOnlyList<string> row = table.Rows[rowIndex];
                if (row.Count < 2) {
                    continue;
                }

                if (NormalizeMarkdownComparison(row[0]) == label &&
                    NormalizeMarkdownComparison(row[row.Count - 1]) == value) {
                    return true;
                }
            }
        }

        return false;
    }

    private static int CompareMarkdownItems(MarkdownItem left, MarkdownItem right) {
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

    private static void AppendBlock(StringBuilder builder, string markdown) {
        if (string.IsNullOrWhiteSpace(markdown)) {
            return;
        }

        if (builder.Length > 0) {
            builder.AppendLine();
            builder.AppendLine();
        }

        builder.Append(markdown.Trim());
    }

    private static string RenderTable(PdfLogicalTable table, PdfLogicalMarkdownOptions options) {
        if (table.Rows.Count == 0) {
            return string.Empty;
        }

        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        if (data.Structure.ColumnCount == 0) {
            return string.Empty;
        }

        var builder = new StringBuilder();
        AppendTableRow(builder, data.Columns, data.Structure.ColumnCount);
        builder.AppendLine();
        AppendTableSeparator(builder, data, options.AlignNumericTableColumns);

        for (int i = 0; i < data.Rows.Count; i++) {
            builder.AppendLine();
            AppendTableRow(builder, data.Rows[i], data.Structure.ColumnCount);
        }

        return builder.ToString();
    }

    private static void AppendTableRow(StringBuilder builder, IReadOnlyList<string> row, int columnCount) {
        builder.Append('|');
        for (int i = 0; i < columnCount; i++) {
            string cell = i < row.Count ? row[i] : string.Empty;
            builder.Append(' ');
            builder.Append(EscapeTableCell(cell));
            builder.Append(" |");
        }
    }

    private static void AppendTableSeparator(StringBuilder builder, PdfLogicalTableData data, bool alignNumericColumns) {
        builder.Append('|');
        for (int i = 0; i < data.Structure.ColumnCount; i++) {
            builder.Append(alignNumericColumns && data.IsNumericColumn(i) ? " ---: |" : " --- |");
        }
    }

    private static string FormatListMarker(string marker) {
        string trimmed = marker.Trim();
        if (trimmed.Length == 0) {
            return "-";
        }

        if (IsNumericMarker(trimmed)) {
            char last = trimmed[trimmed.Length - 1];
            return last == '.' || last == ')'
                ? trimmed
                : trimmed + ".";
        }

        return "-";
    }

    private static bool IsNumericMarker(string marker) {
        int digitCount = 0;
        for (int i = 0; i < marker.Length; i++) {
            char ch = marker[i];
            if (char.IsDigit(ch)) {
                digitCount++;
                continue;
            }

            if (ch != '.' && ch != ')') {
                return false;
            }
        }

        return digitCount > 0;
    }

    private static string EscapeInline(string text) {
        if (string.IsNullOrEmpty(text)) {
            return string.Empty;
        }

        string value = text.Replace("\r", " ").Replace("\n", " ").Trim();
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
                ch == '>') {
                builder.Append('\\');
            }

            builder.Append(ch);
        }

        EscapeLinePrefix(builder);
        return builder.ToString();
    }

    private static string EscapeTableCell(string text) {
        return EscapeInline(text).Replace("|", "\\|");
    }

    private static string EscapeLinkTarget(string uri) {
        return uri.Replace(")", "%29");
    }

    private static string? FormatLinkAnnotationTarget(PdfLogicalLinkAnnotation link) {
        if (!string.IsNullOrEmpty(link.Uri)) {
            return link.Uri;
        }

        if (!string.IsNullOrEmpty(link.DestinationName)) {
            return link.DestinationName;
        }

        if (!string.IsNullOrEmpty(link.NamedAction)) {
            return "named action " + link.NamedAction;
        }

        if (!string.IsNullOrEmpty(link.RemoteFile)) {
            return FormatRemoteLinkAnnotationTarget(link);
        }

        if (!link.DestinationPageNumber.HasValue) {
            return null;
        }

        var builder = new StringBuilder();
        builder.Append("page ");
        builder.Append(link.DestinationPageNumber.Value.ToString(CultureInfo.InvariantCulture));
        if (link.DestinationMode.HasValue) {
            builder.Append(", ");
            builder.Append(link.DestinationMode.Value.ToString());
        }

        AppendCoordinate(builder, "left", link.DestinationLeft);
        AppendCoordinate(builder, "bottom", link.DestinationBottom);
        AppendCoordinate(builder, "right", link.DestinationRight);
        AppendCoordinate(builder, "top", link.DestinationTop);
        return builder.ToString();
    }

    private static string FormatRemoteLinkAnnotationTarget(PdfLogicalLinkAnnotation link) {
        var builder = new StringBuilder();
        builder.Append("remote file ");
        builder.Append(link.RemoteFile);

        if (!string.IsNullOrEmpty(link.RemoteDestinationName)) {
            builder.Append(", destination ");
            builder.Append(link.RemoteDestinationName);
            return builder.ToString();
        }

        if (link.RemoteDestinationPageNumber.HasValue) {
            builder.Append(", page ");
            builder.Append(link.RemoteDestinationPageNumber.Value.ToString(CultureInfo.InvariantCulture));
        }

        if (link.RemoteDestinationMode.HasValue) {
            builder.Append(", ");
            builder.Append(link.RemoteDestinationMode.Value);
        }

        if (link.RemoteDestinationLeft.HasValue) {
            builder.Append(", left ");
            builder.Append(link.RemoteDestinationLeft.Value.ToString("0.###", CultureInfo.InvariantCulture));
        }

        if (link.RemoteDestinationBottom.HasValue) {
            builder.Append(", bottom ");
            builder.Append(link.RemoteDestinationBottom.Value.ToString("0.###", CultureInfo.InvariantCulture));
        }

        if (link.RemoteDestinationRight.HasValue) {
            builder.Append(", right ");
            builder.Append(link.RemoteDestinationRight.Value.ToString("0.###", CultureInfo.InvariantCulture));
        }

        if (link.RemoteDestinationTop.HasValue) {
            builder.Append(", top ");
            builder.Append(link.RemoteDestinationTop.Value.ToString("0.###", CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }

    private static void AppendCoordinate(StringBuilder builder, string name, double? value) {
        if (!value.HasValue) {
            return;
        }

        builder.Append(", ");
        builder.Append(name);
        builder.Append(' ');
        builder.Append(value.Value.ToString("0.###", CultureInfo.InvariantCulture));
    }

    private static bool IsSafeLinkUri(string uri) {
        if (!Guard.IsUriAction(uri)) {
            return false;
        }

        if (!Uri.TryCreate(uri, UriKind.Absolute, out Uri? parsed)) {
            return true;
        }

        return string.Equals(parsed.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(parsed.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(parsed.Scheme, Uri.UriSchemeMailto, StringComparison.OrdinalIgnoreCase);
    }

    private static void EscapeLinePrefix(StringBuilder builder) {
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

    private static string NormalizeMarkdownComparison(string? text) {
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

    private sealed class MarkdownItem {
        public MarkdownItem(double? y, double x, int sequence, string markdown) {
            Y = y;
            X = x;
            Sequence = sequence;
            Markdown = markdown;
        }

        public double? Y { get; }

        public double X { get; }

        public int Sequence { get; }

        public string Markdown { get; }
    }
}
