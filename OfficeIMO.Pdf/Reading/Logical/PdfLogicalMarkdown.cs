using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Options for rendering a logical PDF read model as Markdown.
/// </summary>
public sealed class PdfLogicalMarkdownOptions {
    /// <summary>Emit horizontal-rule separators between logical pages.</summary>
    public bool IncludePageSeparators { get; set; } = true;

    /// <summary>
    /// Emit readable placeholder text for images that are not exported as Markdown images. False by default, so an image
    /// that cannot be carried over never adds text to the document body.
    /// </summary>
    public bool IncludeImagePlaceholders { get; set; }

    /// <summary>
    /// How images are written. The default embeds each image as a base64 data URI; <see cref="PdfLogicalMarkdownImageExportMode.File"/>
    /// writes image files to <see cref="ImageDirectory"/> and references them by file name.
    /// </summary>
    public PdfLogicalMarkdownImageExportMode ImageExportMode { get; set; } = PdfLogicalMarkdownImageExportMode.Base64;

    /// <summary>
    /// Directory that receives image files when <see cref="ImageExportMode"/> is <see cref="PdfLogicalMarkdownImageExportMode.File"/>.
    /// Null uses the current working directory.
    /// </summary>
    public string? ImageDirectory { get; set; }

    /// <summary>Maximum image size embedded as one base64 data URI. Larger images are left out. The default is 32 MiB.</summary>
    public long MaxEmbeddedImageBytes { get; set; } = 32L * 1024L * 1024L;

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
/// How a logical PDF Markdown export writes images.
/// </summary>
public enum PdfLogicalMarkdownImageExportMode {
    /// <summary>Embed images as base64 data URIs.</summary>
    Base64 = 0,

    /// <summary>Write images to <see cref="PdfLogicalMarkdownOptions.ImageDirectory"/> and reference them by relative file name.</summary>
    File = 1,

    /// <summary>Leave images out.</summary>
    None = 2
}

/// <summary>
/// Markdown rendering helpers for the first-party logical PDF read model.
/// </summary>
public static class PdfLogicalMarkdownExtensions {
    /// <summary>
    /// Renders the logical PDF document as Markdown using headings, paragraphs, lists, tables, and optional annotations from the existing logical model.
    /// </summary>
    public static string ExtractMarkdown(this PdfDocumentReadResult document, PdfLogicalMarkdownOptions? options = null) {
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
    public static string ExtractMarkdown(this PdfLogicalPage page, PdfLogicalMarkdownOptions? options = null) {
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
            bool continuesList = i > 0 && items[i].IsListItem && items[i - 1].IsListItem;
            AppendBlock(builder, items[i].Markdown, tight: continuesList && !items[i].RequiresBlankLine,
                preserveIndent: continuesList);
        }
    }

    private static List<MarkdownItem> BuildPageItems(PdfLogicalPage page, PdfLogicalMarkdownOptions options) {
        var items = new List<MarkdownItem>();
        int sequence = 0;
        bool writeImages = options.ImageExportMode != PdfLogicalMarkdownImageExportMode.None || options.IncludeImagePlaceholders;
        HashSet<int>? emittedImages = null;
        Stack<(int Level, int ContentIndent)>? listParents = null;

        IReadOnlyList<PdfLogicalReadingOrderItem> readingOrder =
            PdfLogicalReadingOrderAnalysis.Analyze(page, PdfLogicalReadingOrderScope.PageContent);
        foreach (PdfLogicalReadingOrderItem item in readingOrder) {
            string markdown;
            bool requiresBlankLine = false;
            switch (item.Kind) {
                case PdfLogicalReadingOrderKind.Heading:
                    PdfLogicalHeading heading = page.Headings[item.SourceIndex];
                    int level = Math.Min(Math.Max(heading.Level, 1), 6);
                    markdown = new string('#', level) + " " + EscapeInline(heading.Text, escapeLinePrefix: false);
                    break;
                case PdfLogicalReadingOrderKind.Paragraph:
                    markdown = EscapeInline(page.Paragraphs[item.SourceIndex].Text);
                    break;
                case PdfLogicalReadingOrderKind.ListItem:
                    PdfLogicalListItem listItem = page.ListItems[item.SourceIndex];
                    listParents ??= new Stack<(int Level, int ContentIndent)>();
                    markdown = RenderListItem(listItem, listParents, out requiresBlankLine);
                    break;
                case PdfLogicalReadingOrderKind.Table:
                    markdown = RenderTable(page.Tables[item.SourceIndex], options);
                    break;
                case PdfLogicalReadingOrderKind.TextBlock:
                    markdown = EscapeInline(page.TextBlocks[item.SourceIndex].Text);
                    break;
                case PdfLogicalReadingOrderKind.Image when writeImages &&
                    item.HasGeometry && !item.IsClipped && item.PlacementIndex >= 0:
                    PdfLogicalImage image = page.Images[item.SourceIndex];
                    if (emittedImages?.Contains(item.SourceIndex) == true) continue;
                    markdown = FormatImage(page, item.SourceIndex, image.Placements[item.PlacementIndex], options);
                    if (markdown.Length > 0) (emittedImages ??= new HashSet<int>()).Add(item.SourceIndex);
                    break;
                default: continue;
            }
            if (markdown.Length == 0) continue;
            bool isListItem = item.Kind == PdfLogicalReadingOrderKind.ListItem;
            if (!isListItem) listParents?.Clear();
            items.Add(new MarkdownItem(null, 0, sequence++, markdown, canonical: true,
                isListItem: isListItem, requiresBlankLine: requiresBlankLine));
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

        if (writeImages) {
            for (int i = 0; i < page.Images.Count; i++) {
                if (emittedImages?.Contains(i) == true) continue;
                PdfLogicalImage image = page.Images[i];
                string markdown = string.Empty;
                for (int placementIndex = 0; placementIndex < image.Placements.Count && markdown.Length == 0; placementIndex++) {
                    markdown = FormatImage(page, i, image.Placements[placementIndex], options, allowPlaceholder: false);
                }
                if (markdown.Length == 0 && options.IncludeImagePlaceholders &&
                    (image.Placements.Count == 0 || PdfImagePlacementImportPolicy.HasVisiblePlacement(page, image))) {
                    markdown = FormatImagePlaceholder(image);
                }
                if (markdown.Length > 0) items.Add(new MarkdownItem(null, 0, sequence++, markdown));
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

    /// <summary>
    /// Writes one image placement as a Markdown image. Placements that cannot be carried over safely are left out
    /// (or become a placeholder only when the caller asked for placeholders); clipped placements carry only their
    /// visible pixels.
    /// </summary>
    private static string FormatImage(
        PdfLogicalPage page,
        int imageIndex,
        PdfImagePlacement placement,
        PdfLogicalMarkdownOptions options,
        bool allowPlaceholder = true) {
        PdfLogicalImage image = page.Images[imageIndex];
        PdfImagePlacementImportAssessment assessment = PdfImagePlacementImportPolicy.Analyze(page, image, placement);
        if (assessment.IsSuppressed) return string.Empty;
        if (options.ImageExportMode != PdfLogicalMarkdownImageExportMode.None &&
            assessment.CanImport &&
            page.RotationDegrees % 360 == 0 &&
            placement.A >= 0D && placement.D >= 0D &&
            !assessment.HasNonDefaultOpacity &&
            !assessment.HasNonNormalBlendMode &&
            PdfImagePlacementImportPolicy.TryGetVisibleImagePayload(image, assessment, out byte[] bytes, out string mimeType, out string extension) &&
            OfficeIMO.Drawing.OfficeImageInfo.IsBrowserPreviewSafeContentType(mimeType)) {
            string? source = options.ImageExportMode == PdfLogicalMarkdownImageExportMode.File
                ? WriteImageFile(image, imageIndex, placement, bytes, extension, options)
                : bytes.LongLength <= options.MaxEmbeddedImageBytes
                    ? "data:" + mimeType + ";base64," + Convert.ToBase64String(bytes)
                    : null;
            if (source != null) return "![" + EscapeInline(PdfImagePlacementImportPolicy.GetAlternativeText(image, placement)) + "](" + source + ")";
        }

        return allowPlaceholder && options.IncludeImagePlaceholders ? FormatImagePlaceholder(image) : string.Empty;
    }

    private static string WriteImageFile(PdfLogicalImage image, int imageIndex, PdfImagePlacement placement, byte[] bytes, string extension, PdfLogicalMarkdownOptions options) {
        string directory = options.ImageDirectory ?? Directory.GetCurrentDirectory();
        Directory.CreateDirectory(directory);
        // Page number and image position on the page keep names unique within one export; resource names can
        // collide after sanitizing, so they are not used.
        var name = new StringBuilder("pdf-p");
        name.Append(image.PageNumber.ToString(CultureInfo.InvariantCulture)).Append("-image");
        name.Append((imageIndex + 1).ToString(CultureInfo.InvariantCulture));
        int placementIndex = 0;
        while (placementIndex < image.Placements.Count && !ReferenceEquals(image.Placements[placementIndex], placement)) placementIndex++;
        if (placementIndex > 0) name.Append('-').Append((placementIndex + 1).ToString(CultureInfo.InvariantCulture));
        string normalizedExtension = string.IsNullOrEmpty(extension) ? ".png" : (extension[0] == '.' ? extension : "." + extension);
        string fileName = name.Append(normalizedExtension).ToString();
        File.WriteAllBytes(Path.Combine(directory, fileName), bytes);
        return fileName;
    }

    private static string FormatImagePlaceholder(PdfLogicalImage image) {
        string description = "[Image: page " + image.PageNumber + ", resource " + EscapeInline(image.ResourceName) + ", " + image.Width + "x" + image.Height;
        if (!string.IsNullOrEmpty(image.MimeType)) description += ", " + EscapeInline(image.MimeType!);
        return description + "]";
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
        if (left.Canonical != right.Canonical) return left.Canonical ? -1 : 1;
        if (left.Canonical) return left.Sequence.CompareTo(right.Sequence);
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

    private static void AppendBlock(StringBuilder builder, string markdown, bool tight = false, bool preserveIndent = false) {
        if (string.IsNullOrWhiteSpace(markdown)) {
            return;
        }

        if (builder.Length > 0) {
            builder.AppendLine();
            if (!tight) builder.AppendLine();
        }

        builder.Append(preserveIndent ? markdown.TrimEnd() : markdown.Trim());
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

    private static string RenderListItem(PdfLogicalListItem item, Stack<(int Level, int ContentIndent)> parents,
        out bool requiresBlankLine) {
        int level = Math.Max(item.Level, 1);
        string marker = FormatListMarker(item.Marker);
        bool startsNestedList = parents.Count > 0 && parents.Peek().Level < level;
        // A non-one ordered list needs a blank line after its parent's paragraph to preserve its start number.
        requiresBlankLine = startsNestedList && marker is not ("-" or "1." or "1)");

        while (parents.Count > 0 && parents.Peek().Level >= level) parents.Pop();
        int indent = (level - 1) * 2;
        if (parents.Count > 0) {
            (int parentLevel, int parentContentIndent) = parents.Peek();
            indent = parentContentIndent + (level - parentLevel - 1) * 2;
        }

        parents.Push((level, indent + marker.Length + 1));
        return new string(' ', indent) + marker + " " + (item.IsContentsEntry ? RenderContentsEntry(item) : EscapeInline(item.Text));
    }

    private static string RenderContentsEntry(PdfLogicalListItem item) {
        string? target = item.Link is null ? null : FormatContentsLinkTarget(item.Link);
        // Inside link brackets a title cannot start a block; on its own, "1. Intro" or "# Intro" would.
        string title = EscapeInline(item.Text, escapeLinePrefix: target is null);
        string text = target is null ? title : "[" + title + "](" + target + ")";
        return string.IsNullOrEmpty(item.TargetPageLabel) ? text : text + " \u2014 " + EscapeInline(item.TargetPageLabel!, escapeLinePrefix: false);
    }

    // Named destinations and pages use PDF open-parameter fragments, the same form a viewer accepts after the file name.
    private static string? FormatContentsLinkTarget(PdfLogicalLinkAnnotation link) {
        if (!string.IsNullOrEmpty(link.Uri)) return IsSafeLinkUri(link.Uri!) ? EscapeLinkTarget(link.Uri!) : null;
        if (!string.IsNullOrEmpty(link.DestinationName)) return "#" + EscapeContentsFragment(link.DestinationName!);
        return link.SourceLink.DestinationPageNumber.HasValue
            ? "#page=" + link.SourceLink.DestinationPageNumber.Value.ToString(CultureInfo.InvariantCulture)
            : null;
    }

    private static string EscapeContentsFragment(string name) {
        var builder = new StringBuilder(name.Length);
        foreach (char ch in name) {
            if (char.IsWhiteSpace(ch) || char.IsControl(ch) || ch is '(' or ')' or '<' or '>' or '[' or ']' or '%' or '&' or '\\' or '`') {
                builder.Append(Uri.EscapeDataString(ch.ToString()));
            } else {
                builder.Append(ch);
            }
        }
        return builder.ToString();
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

    private static string EscapeInline(string text, bool escapeLinePrefix = true) {
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

        if (escapeLinePrefix) EscapeLinePrefix(builder);
        return builder.ToString();
    }

    private static string EscapeTableCell(string text) {
        // A cell never starts a Markdown line, so block markers such as "1." or "+2.4" stay literal.
        return EscapeInline(text, escapeLinePrefix: false).Replace("|", "\\|");
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
        bool followedBySpace = index + 1 >= builder.Length || char.IsWhiteSpace(builder[index + 1]);
        // Only escape text that CommonMark would read as a block marker: "# " headings, "- "/"+ " list items,
        // "---" breaks and quotes. "-2.8%" or "+10.6%" are ordinary text.
        if (first == '>' ||
            (first == '#' && (followedBySpace || IsAtxRun(builder, index))) ||
            ((first == '-' || first == '+') && followedBySpace) ||
            (first == '-' && IsThematicBreak(builder, index))) {
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

        if (digitEnd - index <= 9 && digitEnd < builder.Length && (builder[digitEnd] == '.' || builder[digitEnd] == ')') &&
            (digitEnd + 1 >= builder.Length || char.IsWhiteSpace(builder[digitEnd + 1]))) {
            builder.Insert(digitEnd, '\\');
        }
    }

    private static bool IsAtxRun(StringBuilder builder, int index) {
        int end = index;
        while (end < builder.Length && builder[end] == '#') end++;
        return end - index <= 6 && (end >= builder.Length || char.IsWhiteSpace(builder[end]));
    }

    private static bool IsThematicBreak(StringBuilder builder, int index) {
        int dashes = 0;
        for (int position = index; position < builder.Length; position++) {
            char ch = builder[position];
            if (ch == '-') dashes++;
            else if (!char.IsWhiteSpace(ch)) return false;
        }
        return dashes >= 3;
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
        public MarkdownItem(double? y, double x, int sequence, string markdown, bool canonical = false,
            bool isListItem = false, bool requiresBlankLine = false) {
            Y = y;
            X = x;
            Sequence = sequence;
            Markdown = markdown;
            Canonical = canonical;
            IsListItem = isListItem;
            RequiresBlankLine = requiresBlankLine;
        }

        public bool IsListItem { get; }
        public bool RequiresBlankLine { get; }

        public double? Y { get; }

        public double X { get; }

        public int Sequence { get; }

        public string Markdown { get; }

        public bool Canonical { get; }
    }
}
