using System.Globalization;
using System.Text;

namespace OfficeIMO.PowerPoint;

/// <summary>
/// Chunked extraction helpers intended for AI ingestion.
/// </summary>
public static class PowerPointExtractionExtensions {
    /// <summary>
    /// Options controlling PowerPoint extraction behavior.
    /// </summary>
    public sealed class PowerPointExtractOptions {
        /// <summary>
        /// When true, include speaker notes in output. Default: true.
        /// </summary>
        public bool IncludeNotes { get; set; } = true;

        /// <summary>
        /// When true, include slide tables in output. Default: true.
        /// </summary>
        public bool IncludeTables { get; set; } = true;

        /// <summary>
        /// When true, include hidden shapes in output. Default: true.
        /// </summary>
        public bool IncludeHiddenShapes { get; set; } = true;

        /// <summary>When false, projects plain text without building Markdown. Default: true.</summary>
        public bool IncludeMarkdown { get; set; } = true;
    }

    /// <summary>
    /// Extracts a presentation into slide-aligned chunks (one chunk per slide by default).
    /// </summary>
    public static IEnumerable<PowerPointExtractChunk> ExtractMarkdownChunks(
        this PowerPointDocument presentation,
        PowerPointExtractOptions? extract = null,
        PowerPointExtractChunkingOptions? chunking = null,
        string? sourcePath = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(presentation);
        PowerPointExtractOptions snapshot = Snapshot(extract);
        // The options are read now; changing them afterwards does not change the chunks.
        return ExtractDocumentChunks(presentation, snapshot, ResolveMaxChars(chunking), sourcePath, cancellationToken);
    }

    // A presentation whose SDK package is not open yet is read with the streamed slide reader, which gives the same
    // chunks; when it cannot read the package, the model does, with its own errors.
    private static IEnumerable<PowerPointExtractChunk> ExtractDocumentChunks(PowerPointDocument presentation,
        PowerPointExtractOptions extract, int maxChars, string? sourcePath, CancellationToken cancellationToken) {
        List<PowerPointExtractChunk>? streamed = presentation.TryExtractChunksWithoutSdk(extract, maxChars, sourcePath, cancellationToken);
        IEnumerable<PowerPointExtractChunk> chunks = streamed
            ?? ExtractChunksCore(EnumerateModelSlides(presentation, extract, cancellationToken), extract, maxChars, sourcePath, cancellationToken);
        foreach (PowerPointExtractChunk chunk in chunks) yield return chunk;
    }

    internal static IEnumerable<PowerPointExtractChunk> ExtractMarkdownChunks(PowerPointSlideReader reader,
        PowerPointExtractOptions? extract, PowerPointExtractChunkingOptions? chunking, string? sourcePath,
        CancellationToken cancellationToken) {
        PowerPointExtractOptions snapshot = Snapshot(extract);
        return ExtractChunksCore(EnumerateStreamedSlides(reader, snapshot, cancellationToken), snapshot,
            ResolveMaxChars(chunking), sourcePath, cancellationToken);
    }

    /// <summary>Extracts the chunks of one streamed slide; <paramref name="extract"/> must already be snapshotted.</summary>
    internal static IEnumerable<PowerPointExtractChunk> ExtractSlideChunks(PowerPointSlideContent slide,
        PowerPointExtractOptions extract, int maxChars, string? sourcePath, CancellationToken cancellationToken) =>
        ExtractSlideChunks(ToExtractedSlide(slide, extract), extract, maxChars, sourcePath, cancellationToken);

    internal static int ResolveMaxChars(PowerPointExtractChunkingOptions? chunking) => Math.Max(256, chunking?.MaxChars ?? 8000);

    internal static PowerPointExtractOptions Snapshot(PowerPointExtractOptions? extract) => new PowerPointExtractOptions {
        IncludeNotes = extract?.IncludeNotes ?? true,
        IncludeTables = extract?.IncludeTables ?? true,
        IncludeHiddenShapes = extract?.IncludeHiddenShapes ?? true,
        IncludeMarkdown = extract?.IncludeMarkdown ?? true
    };

    private static IEnumerable<ExtractedSlide> EnumerateModelSlides(PowerPointDocument presentation,
        PowerPointExtractOptions extract, CancellationToken cancellationToken) {
        for (int i = 0; i < presentation.Slides.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            PowerPointSlide slide = presentation.Slides[i];
            IEnumerable<PowerPointShapeContent> shapes =
                PowerPointSlideContentProjection.ProjectShapes(slide, extract.IncludeHiddenShapes);
            string? notes = extract.IncludeNotes && slide.Notes.TryGetExistingText(out string text) ? text : null;
            yield return new ExtractedSlide(i + 1, shapes, notes);
        }
    }

    private static IEnumerable<ExtractedSlide> EnumerateStreamedSlides(PowerPointSlideReader reader,
        PowerPointExtractOptions extract, CancellationToken cancellationToken) {
        foreach (PowerPointSlideContent slide in reader.ReadSlides(extract.IncludeNotes, cancellationToken)) {
            yield return ToExtractedSlide(slide, extract);
        }
    }

    private static ExtractedSlide ToExtractedSlide(PowerPointSlideContent slide, PowerPointExtractOptions extract) {
        IEnumerable<PowerPointShapeContent> shapes = extract.IncludeHiddenShapes
            ? slide.Shapes
            : slide.Shapes.Where(static shape => shape.IsVisible);
        return new ExtractedSlide(slide.Number, shapes, extract.IncludeNotes ? slide.Notes : null);
    }

    private static IEnumerable<PowerPointExtractChunk> ExtractChunksCore(IEnumerable<ExtractedSlide> slides,
        PowerPointExtractOptions extract, int maxChars, string? sourcePath, CancellationToken cancellationToken) {
        foreach (ExtractedSlide slide in slides) {
            foreach (PowerPointExtractChunk chunk in ExtractSlideChunks(slide, extract, maxChars, sourcePath, cancellationToken)) {
                yield return chunk;
            }
        }
    }

    private static IEnumerable<PowerPointExtractChunk> ExtractSlideChunks(ExtractedSlide slide,
        PowerPointExtractOptions extract, int maxChars, string? sourcePath, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();

        int slideNumber = slide.Number;
        int i = slideNumber - 1;
        List<string>? warnings = null;

        if (!extract.IncludeMarkdown) {
            string plain = BuildSlidePlainText(slide, extract, cancellationToken);
            for (int offset = 0, fragment = 0; offset < plain.Length; fragment++) {
                cancellationToken.ThrowIfCancellationRequested();
                int length = Math.Min(maxChars, plain.Length - offset);
                if (offset + length < plain.Length && char.IsHighSurrogate(plain[offset + length - 1]) && char.IsLowSurrogate(plain[offset + length])) length--;
                yield return new PowerPointExtractChunk {
                    Id = BuildStableId("ppt-text", sourcePath, slideNumber) + ":" + fragment.ToString(CultureInfo.InvariantCulture),
                    Location = new PowerPointExtractLocation { Path = sourcePath, Slide = slideNumber, BlockIndex = i },
                    Text = plain.Substring(offset, length), Markdown = null
                };
                offset += length;
            }
            yield break;
        }

        var md = new StringBuilder();
        md.Append("## Slide ").AppendLine(slideNumber.ToString(System.Globalization.CultureInfo.InvariantCulture));
        md.AppendLine();

        int tableIndex = 0;
        foreach (PowerPointShapeContent shape in slide.Shapes) {
            cancellationToken.ThrowIfCancellationRequested();

            if (shape.Kind == PowerPointShapeContentType.TextBox) {
                AppendTextBoxMarkdown(md, shape);
            } else if (extract.IncludeTables && shape.Table != null) {
                AppendTableMarkdown(md, shape.Table, ref tableIndex);
            }
        }

        if (slide.Notes != null) {
            md.AppendLine("### Notes");
            md.AppendLine();
            md.AppendLine(slide.Notes);
            md.AppendLine();
        }

        var markdown = NormalizeLineEndings(md.ToString()).TrimEnd();
        if (markdown.Length > maxChars) {
            int truncationLength = maxChars;
            if (truncationLength > 0
                && char.IsHighSurrogate(markdown[truncationLength - 1])
                && char.IsLowSurrogate(markdown[truncationLength])) {
                truncationLength--;
            }
            markdown = markdown.Substring(0, truncationLength)
                + "\n\n<!-- truncated -->";
            warnings = new List<string> { "Markdown truncated to MaxChars." };
        }

        var id = BuildStableId("ppt-md", sourcePath, slideNumber);
        yield return new PowerPointExtractChunk {
            Id = id,
            Location = new PowerPointExtractLocation {
                Path = sourcePath,
                Slide = slideNumber,
                BlockIndex = i
            },
            Text = markdown,
            Markdown = markdown,
            Warnings = warnings
        };
    }

    private static string NormalizeLineEndings(string value) =>
        value.Replace("\r\n", "\n").Replace('\r', '\n');

    private static string BuildSlidePlainText(ExtractedSlide slide, PowerPointExtractOptions options, CancellationToken token) {
        var text = new StringBuilder();
        foreach (PowerPointShapeContent shape in slide.Shapes) {
            token.ThrowIfCancellationRequested();
            if (shape.Kind == PowerPointShapeContentType.TextBox) {
                foreach (PowerPointParagraphContent paragraph in shape.Paragraphs) text.AppendLine(RenderParagraphText(paragraph, renderMarkdown: false));
            } else if (options.IncludeTables && shape.Table != null) {
                foreach (PowerPointTableRowContent row in shape.Table.Rows) {
                    token.ThrowIfCancellationRequested();
                    text.AppendLine(string.Join("\t", row.Cells.Select(cell => string.Join(" ", cell.Paragraphs.Select(paragraph => RenderParagraphText(paragraph, renderMarkdown: false))))));
                }
            }
        }
        if (slide.Notes != null) text.AppendLine(slide.Notes);
        return NormalizeLineEndings(text.ToString()).TrimEnd();
    }

    private static void AppendTextBoxMarkdown(StringBuilder markdown, PowerPointShapeContent textBox) {
        bool isTitle = textBox.IsTitle;
        var numberingState = new Dictionary<int, int>();
        var listContentIndents = new Dictionary<int, int>();
        bool appended = false;
        ParagraphMarkdownKind? previousKind = null;
        foreach (PowerPointParagraphContent paragraph in textBox.Paragraphs) {
            string text = RenderParagraphText(paragraph).Trim();
            if (text.Length == 0) {
                numberingState.Clear();
                listContentIndents.Clear();
                if (appended) previousKind = ParagraphMarkdownKind.Plain;
                continue;
            }

            ParagraphMarkdownKind kind = isTitle
                ? ParagraphMarkdownKind.Heading
                : paragraph.IsNumbered
                    ? ParagraphMarkdownKind.NumberedList
                    : !string.IsNullOrEmpty(paragraph.BulletCharacter)
                        ? ParagraphMarkdownKind.BulletList
                        : ParagraphMarkdownKind.Plain;
            if (appended && previousKind.HasValue
                && (!IsListKind(previousKind.Value) || !IsListKind(kind))) {
                markdown.AppendLine();
            }

            if (kind == ParagraphMarkdownKind.Heading) {
                numberingState.Clear();
                listContentIndents.Clear();
                markdown.Append("### ").AppendLine(text);
            } else {
                // DrawingML list levels are defined only for levels 0-8. Clamp untrusted
                // package values before they can drive indentation allocation.
                int level = Math.Min(8, Math.Max(0, paragraph.Level ?? 0));
                if (kind == ParagraphMarkdownKind.NumberedList) {
                    RemoveListState(numberingState, level,
                        includeLevel: false);
                    RemoveListState(listContentIndents, level,
                        includeLevel: false);
                } else if (kind == ParagraphMarkdownKind.BulletList) {
                    RemoveListState(numberingState, level,
                        includeLevel: true);
                    RemoveListState(listContentIndents, level,
                        includeLevel: false);
                } else {
                    numberingState.Clear();
                    listContentIndents.Clear();
                }
                int indent = ResolveListIndent(level, listContentIndents);
                if (kind == ParagraphMarkdownKind.NumberedList) {
                    int number = ResolveNumberingValue(paragraph, level, numberingState);
                    string markdownMarker = number.ToString(CultureInfo.InvariantCulture) + ".";
                    string powerPointMarker = PowerPointNumberingFormatter.FormatMarker(
                        number, paragraph.NumberingScheme);
                    markdown.Append(' ', indent);
                    markdown.Append(markdownMarker).Append(' ');
                    // CommonMark only recognizes decimal ordered-list markers.
                    // Retain the displayed PowerPoint marker as visible text.
                    if (!string.Equals(markdownMarker, powerPointMarker, StringComparison.Ordinal)) {
                        markdown.Append(powerPointMarker).Append(' ');
                    }
                    markdown.AppendLine(text);
                    UpdateListContentIndent(listContentIndents, level,
                        checked(indent + markdownMarker.Length + 1));
                } else if (kind == ParagraphMarkdownKind.BulletList) {
                    char markdownMarker = ResolveBulletMarker(
                        paragraph.BulletCharacter);
                    markdown.Append(' ', indent);
                    markdown.Append(markdownMarker).Append(' ');
                    if (!string.Equals(markdownMarker.ToString(),
                            paragraph.BulletCharacter,
                            StringComparison.Ordinal)) {
                        markdown.Append(paragraph.BulletCharacter)
                            .Append(' ');
                    }
                    markdown.AppendLine(text);
                    UpdateListContentIndent(listContentIndents, level,
                        checked(indent + 2));
                } else {
                    markdown.AppendLine(text);
                }
            }
            appended = true;
            previousKind = kind;
        }

        if (appended) {
            markdown.AppendLine();
        }
    }

    private static int ResolveNumberingValue(PowerPointParagraphContent paragraph, int level,
        IDictionary<int, int> numberingState) {
        int number = paragraph.NumberingStartAt
            ?? (numberingState.TryGetValue(level, out int previous) ? previous + 1 : 1);
        numberingState[level] = number;
        return number;
    }

    private static void RemoveListState(IDictionary<int, int> state,
        int level, bool includeLevel) {
        foreach (int key in state.Keys.Where(key => includeLevel
                     ? key >= level
                     : key > level).ToArray()) {
            state.Remove(key);
        }
    }

    private static bool IsListKind(ParagraphMarkdownKind kind) =>
        kind == ParagraphMarkdownKind.BulletList || kind == ParagraphMarkdownKind.NumberedList;

    private static string RenderParagraphText(PowerPointParagraphContent paragraph, bool renderMarkdown = true) {
        var text = new StringBuilder();
        foreach (PowerPointInlineContent inline in paragraph.Inlines) {
            switch (inline.Kind) {
                case PowerPointParagraphInlineKind.Run:
                    if (renderMarkdown && inline.Hyperlink != null && inline.Text.Length > 0) {
                        text.Append('[').Append(EscapeMarkdownLinkLabel(inline.Text)).Append("](<")
                            .Append(EscapeMarkdownLinkDestination(inline.Hyperlink.ToString())).Append(">)");
                    } else {
                        text.Append(inline.Text);
                    }
                    break;
                case PowerPointParagraphInlineKind.LineBreak:
                    text.Append(renderMarkdown ? "<br />" : "\n");
                    break;
                case PowerPointParagraphInlineKind.Field:
                    // A field without its own text element renders the text of its other children.
                    text.Append(inline.FieldFallbackText ?? inline.Text);
                    break;
            }
        }

        return text.Length == 0 ? paragraph.Text : text.ToString();
    }

    private static char ResolveBulletMarker(string? bulletCharacter) {
        return bulletCharacter is "*" or "+" or "-" ? bulletCharacter[0] : '-';
    }

    private static int ResolveListIndent(int level,
        IReadOnlyDictionary<int, int> listContentIndents) {
        if (level <= 0) return 0;
        int conventionalIndent = checked(level * 4);
        return listContentIndents.TryGetValue(level - 1,
                out int parentContentIndent)
            ? Math.Max(conventionalIndent, parentContentIndent)
            : conventionalIndent;
    }

    private static void UpdateListContentIndent(
        IDictionary<int, int> listContentIndents, int level,
        int contentIndent) {
        foreach (int deeperLevel in listContentIndents.Keys
                     .Where(existingLevel => existingLevel > level)
                     .ToArray()) {
            listContentIndents.Remove(deeperLevel);
        }
        listContentIndents[level] = contentIndent;
    }

    private static string EscapeMarkdownLinkLabel(string value) {
        return value.Replace("\\", "\\\\")
            .Replace("[", "\\[")
            .Replace("]", "\\]");
    }

    private static string EscapeMarkdownLinkDestination(string value) {
        var escaped = new StringBuilder(value.Length);
        foreach (char character in value) {
            if (char.IsWhiteSpace(character) || char.IsControl(character)
                || character is '<' or '>' or '\\') {
                escaped.Append(Uri.EscapeDataString(character.ToString()));
            } else {
                escaped.Append(character);
            }
        }
        return escaped.ToString();
    }

    private static void AppendTableMarkdown(StringBuilder markdown, PowerPointTableContent table, ref int tableIndex) {
        List<IReadOnlyList<string>> rows = table.Rows
            .Select(row => (IReadOnlyList<string>)row.Cells.Select(RenderTableCellText).ToList())
            .ToList();

        if (rows.Count == 0 || rows.All(row => row.All(string.IsNullOrWhiteSpace))) {
            return;
        }

        int columnCount = rows.Max(row => row.Count);
        if (columnCount == 0) {
            return;
        }

        tableIndex++;
        markdown.Append("### Table ").AppendLine(tableIndex.ToString(CultureInfo.InvariantCulture));
        markdown.AppendLine();
        AppendMarkdownTable(markdown, rows, columnCount);
        markdown.AppendLine();
    }

    private static string RenderTableCellText(PowerPointTableCellContent cell) {
        string[] paragraphs = cell.Paragraphs
            .Select(paragraph => RenderParagraphText(paragraph).Trim())
            .Where(text => text.Length > 0)
            .ToArray();
        return paragraphs.Length == 0
            ? NormalizeText(cell.Text)
            : string.Join("<br />", paragraphs);
    }

    private static void AppendMarkdownTable(StringBuilder markdown, IReadOnlyList<IReadOnlyList<string>> rows, int columnCount) {
        IReadOnlyList<string> header = rows[0];
        markdown.Append('|');
        for (int i = 0; i < columnCount; i++) {
            markdown.Append(' ').Append(EscapeMarkdownTableCell(GetCellValue(header, i, $"Column {i + 1}"))).Append(" |");
        }
        markdown.AppendLine();

        markdown.Append('|');
        for (int i = 0; i < columnCount; i++) {
            markdown.Append(" --- |");
        }
        markdown.AppendLine();

        foreach (IReadOnlyList<string> row in rows.Skip(1)) {
            markdown.Append('|');
            for (int i = 0; i < columnCount; i++) {
                markdown.Append(' ').Append(EscapeMarkdownTableCell(GetCellValue(row, i, string.Empty))).Append(" |");
            }
            markdown.AppendLine();
        }
    }

    private static string GetCellValue(IReadOnlyList<string> row, int index, string fallback) {
        if (index >= row.Count) {
            return fallback;
        }

        string value = row[index];
        return string.IsNullOrWhiteSpace(value) ? fallback : value;
    }

    private static string EscapeMarkdownTableCell(string value) {
        return (value ?? string.Empty)
            .Replace("\r\n", " ")
            .Replace('\n', ' ')
            .Replace('\r', ' ')
            .Replace("|", "\\|");
    }

    private static string NormalizeText(string? text) {
        return (text ?? string.Empty).Trim();
    }

    private static string BuildStableId(string kind, string? path, int slideNumber) {
        var safe = string.IsNullOrWhiteSpace(path) ? "memory" : System.IO.Path.GetFileName(path!.Trim());
        return $"{kind}:{safe}:s{slideNumber}";
    }

    private readonly struct ExtractedSlide {
        internal ExtractedSlide(int number, IEnumerable<PowerPointShapeContent> shapes, string? notes) {
            Number = number;
            Shapes = shapes;
            Notes = notes;
        }

        internal int Number { get; }
        internal IEnumerable<PowerPointShapeContent> Shapes { get; }
        internal string? Notes { get; }
    }

    private enum ParagraphMarkdownKind {
        Plain,
        Heading,
        BulletList,
        NumberedList
    }
}
