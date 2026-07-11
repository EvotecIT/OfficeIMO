namespace OfficeIMO.AsciiDoc.Markdown;

/// <summary>Canonical, loss-aware Markdown-to-AsciiDoc conversion.</summary>
public static class MarkdownToAsciiDocConverter {
    /// <summary>Converts the OfficeIMO Markdown semantic model to dependency-free AsciiDoc.</summary>
    public static MarkdownAsciiDocConversionResult Convert(
        MarkdownDoc document,
        MarkdownToAsciiDocOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        options ??= new MarkdownToAsciiDocOptions();
        ValidateOptions(options);

        var diagnostics = new List<MarkdownAsciiDocConversionDiagnostic>();
        var blocks = new List<string>();
        if (document.DocumentHeader != null) {
            string attributes = ConvertFrontMatter(document.DocumentHeader, options, diagnostics);
            if (attributes.Length > 0) blocks.Add(attributes);
        }

        bool documentTitleWritten = false;
        for (int index = 0; index < document.Blocks.Count; index++) {
            IMarkdownBlock block = document.Blocks[index];
            string converted = ConvertBlock(block, options, diagnostics, ref documentTitleWritten);
            if (converted.Length > 0) blocks.Add(converted);
        }

        string source = string.Join(options.LineEnding + options.LineEnding, blocks);
        if (source.Length > 0) source += options.LineEnding;
        AsciiDocDocument parsed = AsciiDocDocument.Parse(source).Document;
        return new MarkdownAsciiDocConversionResult(source, parsed, diagnostics);
    }

    private static string ConvertBlock(
        IMarkdownBlock block,
        MarkdownToAsciiDocOptions options,
        List<MarkdownAsciiDocConversionDiagnostic> diagnostics,
        ref bool documentTitleWritten) {
        switch (block) {
            case HeadingBlock heading:
                bool documentTitle = options.FirstLevelOneHeadingIsDocumentTitle && !documentTitleWritten && heading.Level == 1;
                if (documentTitle) documentTitleWritten = true;
                string headingMetadata = ConvertMetadata(heading, null, options.LineEnding);
                int markerLength = documentTitle ? 1 : Math.Min(6, Math.Max(2, heading.Level));
                return headingMetadata + new string('=', markerLength) + " " +
                    MarkdownInlineToAsciiDocConverter.Convert(heading.Inlines, diagnostics, heading);
            case ParagraphBlock paragraph:
                return ConvertMetadata(paragraph, null, options.LineEnding) +
                    MarkdownInlineToAsciiDocConverter.Convert(paragraph.Inlines, diagnostics, paragraph);
            case UnorderedListBlock unordered:
                return ConvertList(unordered.Items, false, unordered, options, diagnostics, ref documentTitleWritten);
            case OrderedListBlock ordered:
                return ConvertList(ordered.Items, true, ordered, options, diagnostics, ref documentTitleWritten);
            case DefinitionListBlock definitions:
                return ConvertDefinitionList(definitions, options, diagnostics, ref documentTitleWritten);
            case CodeBlock code:
                return ConvertCode(code, options, diagnostics);
            case SemanticFencedBlock semantic when string.Equals(semantic.SemanticKind, MarkdownSemanticKinds.Math, StringComparison.OrdinalIgnoreCase):
                string math = Normalize(semantic.Content, options.LineEnding);
                string mathDelimiter = ChooseDelimiter(math, '+');
                return ConvertMetadata(semantic, semantic.Caption, options.LineEnding) +
                    "[stem]" + options.LineEnding + mathDelimiter + options.LineEnding +
                    math + EnsureEnding(math, options.LineEnding) + mathDelimiter;
            case TableBlock table:
                return ConvertTable(table, options, diagnostics, ref documentTitleWritten);
            case CalloutBlock callout:
                return ConvertCallout(callout, options, diagnostics, ref documentTitleWritten);
            case QuoteBlock quote:
                return ConvertQuote(quote, options, diagnostics, ref documentTitleWritten);
            case ImageBlock image:
                return ConvertImage(image, options);
            default:
                return ConvertUnsupported(block, options, diagnostics);
        }
    }

    private static string ConvertList(
        IReadOnlyList<ListItem> items,
        bool ordered,
        MarkdownObject owner,
        MarkdownToAsciiDocOptions options,
        List<MarkdownAsciiDocConversionDiagnostic> diagnostics,
        ref bool documentTitleWritten) {
        var output = new System.Text.StringBuilder();
        output.Append(ConvertMetadata(owner, null, options.LineEnding));
        for (int index = 0; index < items.Count; index++) {
            ListItem item = items[index];
            if (index > 0) output.Append(options.LineEnding);
            string marker = new string(ordered ? '.' : '*', Math.Max(1, item.Level + 1));
            output.Append(marker).Append(' ');
            if (item.IsTask) output.Append(item.Checked ? "[x] " : "[ ] ");
            output.Append(MarkdownInlineToAsciiDocConverter.Convert(item.Content, diagnostics, item));
            for (int childIndex = 0; childIndex < item.Children.Count; childIndex++) {
                output.Append(options.LineEnding).Append('+').Append(options.LineEnding);
                output.Append(ConvertBlock(item.Children[childIndex], options, diagnostics, ref documentTitleWritten));
            }
        }
        return output.ToString();
    }

    private static string ConvertDefinitionList(
        DefinitionListBlock source,
        MarkdownToAsciiDocOptions options,
        List<MarkdownAsciiDocConversionDiagnostic> diagnostics,
        ref bool documentTitleWritten) {
        var output = new System.Text.StringBuilder();
        output.Append(ConvertMetadata(source, null, options.LineEnding));
        for (int index = 0; index < source.Entries.Count; index++) {
            DefinitionListEntry entry = source.Entries[index];
            if (index > 0) output.Append(options.LineEnding);
            output.Append(MarkdownInlineToAsciiDocConverter.Convert(entry.Term, diagnostics, entry)).Append("::");
            if (entry.DefinitionBlocks.Count == 1 && entry.DefinitionBlocks[0] is ParagraphBlock paragraph) {
                output.Append(' ').Append(MarkdownInlineToAsciiDocConverter.Convert(paragraph.Inlines, diagnostics, paragraph));
            } else {
                for (int blockIndex = 0; blockIndex < entry.DefinitionBlocks.Count; blockIndex++) {
                    output.Append(options.LineEnding).Append('+').Append(options.LineEnding);
                    output.Append(ConvertBlock(entry.DefinitionBlocks[blockIndex], options, diagnostics, ref documentTitleWritten));
                }
            }
        }
        return output.ToString();
    }

    private static string ConvertCode(
        CodeBlock source,
        MarkdownToAsciiDocOptions options,
        List<MarkdownAsciiDocConversionDiagnostic> diagnostics) {
        string metadata = ConvertMetadata(source, source.Caption, options.LineEnding);
        var attributes = new List<string> { "source" };
        if (!string.IsNullOrWhiteSpace(source.Language)) attributes.Add(EscapeAttribute(source.Language));
        string content = Normalize(source.Content, options.LineEnding);
        string delimiter = ChooseDelimiter(content, '-');
        return metadata + "[" + string.Join(",", attributes) + "]" + options.LineEnding +
            delimiter + options.LineEnding + content + EnsureEnding(content, options.LineEnding) + delimiter;
    }

    private static string ConvertTable(
        TableBlock source,
        MarkdownToAsciiDocOptions options,
        List<MarkdownAsciiDocConversionDiagnostic> diagnostics,
        ref bool documentTitleWritten) {
        TableRow[] rows = source.EnumerateRows().ToArray();
        int columns = rows.Length == 0
            ? 1
            : Math.Max(1, rows.Max(static row => row.Cells.Sum(static cell => Math.Max(1, cell.ColumnSpan))));
        bool header = rows.Length > 0 && rows[0].IsHeader;
        var attributes = new List<string> { "cols=" + columns + "*" };
        if (header) attributes.Add("%header");
        var output = new System.Text.StringBuilder();
        output.Append(ConvertMetadata(source, null, options.LineEnding));
        output.Append('[').Append(string.Join(",", attributes)).Append(']').Append(options.LineEnding);
        output.Append("|===").Append(options.LineEnding);
        for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++) {
            TableRow row = rows[rowIndex];
            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++) {
                TableCell cell = row.Cells[cellIndex];
                string specifier = BuildCellSpecifier(cell);
                string content = ConvertTableCell(cell, options, diagnostics, ref documentTitleWritten).Replace("|", "\\|");
                output.Append(specifier).Append('|').Append(content).Append(options.LineEnding);
            }
            if (rowIndex + 1 < rows.Length) output.Append(options.LineEnding);
        }
        output.Append("|===");
        return output.ToString();
    }

    private static string ConvertTableCell(
        TableCell cell,
        MarkdownToAsciiDocOptions options,
        List<MarkdownAsciiDocConversionDiagnostic> diagnostics,
        ref bool documentTitleWritten) {
        if (cell.Blocks.Count == 1 && cell.Blocks[0] is ParagraphBlock paragraph) {
            return MarkdownInlineToAsciiDocConverter.Convert(paragraph.Inlines, diagnostics, paragraph);
        }
        var blocks = new List<string>();
        for (int index = 0; index < cell.Blocks.Count; index++) {
            blocks.Add(ConvertBlock(cell.Blocks[index], options, diagnostics, ref documentTitleWritten));
        }
        return string.Join(options.LineEnding + options.LineEnding, blocks);
    }

    private static string BuildCellSpecifier(TableCell cell) {
        var output = new System.Text.StringBuilder();
        if (cell.ColumnSpan > 1 || cell.RowSpan > 1) {
            if (cell.ColumnSpan > 1) output.Append(cell.ColumnSpan);
            if (cell.RowSpan > 1) output.Append('.').Append(cell.RowSpan);
            output.Append('+');
        }
        if (cell.Bold) output.Append('s');
        else if (cell.Italic) output.Append('e');
        else if (cell.Blocks.Count > 1) output.Append('a');
        return output.ToString();
    }

    private static string ConvertCallout(
        CalloutBlock source,
        MarkdownToAsciiDocOptions options,
        List<MarkdownAsciiDocConversionDiagnostic> diagnostics,
        ref bool documentTitleWritten) {
        string kind = NormalizeAdmonitionKind(source.Kind);
        if (string.IsNullOrWhiteSpace(source.Title) && source.ChildBlocks.Count == 1 && source.ChildBlocks[0] is ParagraphBlock paragraph) {
            return ConvertMetadata(source, null, options.LineEnding) + kind + ": " +
                MarkdownInlineToAsciiDocConverter.Convert(paragraph.Inlines, diagnostics, paragraph);
        }
        var body = new List<string>();
        for (int index = 0; index < source.ChildBlocks.Count; index++) {
            body.Add(ConvertBlock(source.ChildBlocks[index], options, diagnostics, ref documentTitleWritten));
        }
        string content = string.Join(options.LineEnding + options.LineEnding, body);
        string delimiter = ChooseDelimiter(content, '=');
        string title = string.IsNullOrWhiteSpace(source.Title) ? string.Empty : "." + source.Title + options.LineEnding;
        return ConvertMetadata(source, null, options.LineEnding) + title + "[" + kind + "]" + options.LineEnding +
            delimiter + options.LineEnding + content + options.LineEnding + delimiter;
    }

    private static string ConvertQuote(
        QuoteBlock source,
        MarkdownToAsciiDocOptions options,
        List<MarkdownAsciiDocConversionDiagnostic> diagnostics,
        ref bool documentTitleWritten) {
        string body;
        if (source.Children.Count > 0) {
            var blocks = new List<string>();
            for (int index = 0; index < source.Children.Count; index++) {
                blocks.Add(ConvertBlock(source.Children[index], options, diagnostics, ref documentTitleWritten));
            }
            body = string.Join(options.LineEnding + options.LineEnding, blocks);
        } else {
            body = string.Join(options.LineEnding, source.Lines);
        }
        string delimiter = ChooseDelimiter(body, '_');
        return ConvertMetadata(source, null, options.LineEnding) + delimiter + options.LineEnding + body + options.LineEnding + delimiter;
    }

    private static string ConvertImage(ImageBlock source, MarkdownToAsciiDocOptions options) {
        string metadata = ConvertMetadata(source, source.Caption, options.LineEnding);
        return metadata + "image::" + source.Path.Replace("[", "\\[").Replace("]", "\\]") +
            "[" + EscapeAttribute(source.Alt ?? string.Empty) +
            (string.IsNullOrEmpty(source.Title) ? string.Empty : ",\"" + EscapeAttribute(source.Title!) + "\"") + "]";
    }

    private static string ConvertUnsupported(
        IMarkdownBlock source,
        MarkdownToAsciiDocOptions options,
        List<MarkdownAsciiDocConversionDiagnostic> diagnostics) {
        MarkdownObject? markdownObject = source as MarkdownObject;
        if (!options.PreserveUnsupportedAsSource) {
            diagnostics.Add(new MarkdownAsciiDocConversionDiagnostic(
                "MDADOC099", AsciiDocMarkdownDiagnosticSeverity.Warning, AsciiDocMarkdownConversionOutcome.Omitted,
                source.GetType().Name, "Unsupported Markdown block omitted by conversion options.", markdownObject?.SourceSpan));
            return string.Empty;
        }
        string markdown = Normalize(source.RenderMarkdown(), options.LineEnding);
        string delimiter = ChooseDelimiter(markdown, '-');
        diagnostics.Add(new MarkdownAsciiDocConversionDiagnostic(
            "MDADOC098", AsciiDocMarkdownDiagnosticSeverity.Warning, AsciiDocMarkdownConversionOutcome.SourceFallback,
            source.GetType().Name, "Unsupported Markdown block retained in a Markdown source listing.", markdownObject?.SourceSpan));
        return "[source,markdown]" + options.LineEnding + delimiter + options.LineEnding + markdown +
            EnsureEnding(markdown, options.LineEnding) + delimiter;
    }

    private static string ConvertFrontMatter(
        FrontMatterBlock source,
        MarkdownToAsciiDocOptions options,
        List<MarkdownAsciiDocConversionDiagnostic> diagnostics) {
        var output = new System.Text.StringBuilder();
        for (int index = 0; index < source.Entries.Count; index++) {
            FrontMatterBlock.Entry entry = source.Entries[index];
            if (index > 0) output.Append(options.LineEnding);
            string value = System.Convert.ToString(entry.Value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
            if (value.IndexOf('\r') >= 0 || value.IndexOf('\n') >= 0) {
                value = value.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ');
                diagnostics.Add(new MarkdownAsciiDocConversionDiagnostic(
                    "MDADOC001", AsciiDocMarkdownDiagnosticSeverity.Warning, AsciiDocMarkdownConversionOutcome.Simplified,
                    "front-matter-multiline", "Multiline front-matter value flattened to one document attribute line.", entry.SourceSpan));
            }
            output.Append(':').Append(entry.Key.Replace(":", "-")).Append(':');
            if (value.Length > 0) output.Append(' ').Append(value);
        }
        return output.ToString();
    }

    private static string ConvertMetadata(MarkdownObject source, string? caption, string lineEnding) {
        var output = new System.Text.StringBuilder();
        if (!string.IsNullOrWhiteSpace(caption)) output.Append('.').Append(caption).Append(lineEnding);
        if (!string.IsNullOrWhiteSpace(source.Attributes.ElementId)) {
            output.Append("[[").Append(source.Attributes.ElementId).Append("]]").Append(lineEnding);
        }
        var values = new List<string>();
        values.AddRange(source.Attributes.Classes.Select(static role => "." + role));
        values.AddRange(source.Attributes.Attributes.Select(static pair => pair.Key + "=\"" + EscapeAttribute(pair.Value ?? string.Empty) + "\""));
        if (values.Count > 0) output.Append('[').Append(string.Join(",", values)).Append(']').Append(lineEnding);
        return output.ToString();
    }

    private static string NormalizeAdmonitionKind(string kind) {
        string value = (kind ?? string.Empty).Trim().ToUpperInvariant();
        switch (value) {
            case "TIP":
            case "IMPORTANT":
            case "WARNING":
            case "CAUTION": return value;
            default: return "NOTE";
        }
    }

    private static string EscapeAttribute(string value) => value.Replace("\\", "\\\\").Replace("]", "\\]").Replace(",", "\\,").Replace("\"", "\\\"");

    private static string Normalize(string value, string lineEnding) =>
        value.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", lineEnding);

    private static string EnsureEnding(string value, string lineEnding) =>
        value.Length == 0 || value.EndsWith("\n", StringComparison.Ordinal) || value.EndsWith("\r", StringComparison.Ordinal)
            ? string.Empty
            : lineEnding;

    private static string ChooseDelimiter(string value, char marker) {
        int length = 4;
        string[] lines = Normalize(value, "\n").Split('\n');
        for (int index = 0; index < lines.Length; index++) {
            string line = lines[index];
            if (line.Length < length || line.Any(character => character != marker)) continue;
            length = line.Length + 1;
        }
        return new string(marker, length);
    }

    private static void Report(
        List<MarkdownAsciiDocConversionDiagnostic> diagnostics,
        MarkdownObject owner,
        string code,
        string feature,
        string message) {
        diagnostics.Add(new MarkdownAsciiDocConversionDiagnostic(
            code,
            AsciiDocMarkdownDiagnosticSeverity.Warning,
            AsciiDocMarkdownConversionOutcome.Simplified,
            feature,
            message,
            owner.SourceSpan));
    }

    private static void ValidateOptions(MarkdownToAsciiDocOptions options) {
        if (options.LineEnding != "\n" && options.LineEnding != "\r" && options.LineEnding != "\r\n") {
            throw new ArgumentException("LineEnding must be LF, CR, or CRLF.", nameof(options));
        }
    }
}
