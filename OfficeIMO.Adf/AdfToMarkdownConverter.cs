using System.IO;
using System.Text;
using OfficeIMO.Markdown;

namespace OfficeIMO.Adf;

internal static partial class AdfToMarkdownConverter {
    internal static string Convert(AdfDocument document, AdfConversionOptions options, List<AdfConversionDiagnostic> diagnostics) {
        var builder = new StringBuilder();
        if (document.ExtensionData.Count > 0) {
            diagnostics.Add(Warning("ADF_ROOT_PROPERTIES_DROPPED", "$", "ADF root extension properties cannot be represented in Markdown and were omitted."));
        }
        for (int i = 0; i < document.Content.Count; i++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            AppendBlock(builder, document.Content[i], "$.content[" + i + "]", options, diagnostics, 0);
            if (i < document.Content.Count - 1 && !EndsWithBlankLine(builder)) builder.AppendLine().AppendLine();
            if (builder.Length > options.MaxOutputCharacters) throw AdfGraphGuard.Limit("MaxOutputCharacters");
        }
        return builder.ToString();
    }

    private static void AppendBlock(StringBuilder builder, AdfNode node, string path, AdfConversionOptions options, List<AdfConversionDiagnostic> diagnostics, int listDepth) {
        options.CancellationToken.ThrowIfCancellationRequested();
        ReportBlockProperties(node, path, diagnostics);
        switch (node.Type) {
            case "paragraph":
                if (node.Content.Count == 0) {
                    diagnostics.Add(Warning("ADF_EMPTY_PARAGRAPH_DROPPED", path, "Markdown cannot preserve an explicit empty ADF paragraph, so the paragraph was omitted."));
                }
                builder.Append(RenderInlineContent(node, path, diagnostics));
                break;
            case "heading":
                int level = node.GetInt32Attribute("level") ?? 1;
                level = Math.Max(1, Math.Min(6, level));
                if (node.ExtensionData.Count > 0 || node.Attributes.Keys.Any(key => !string.Equals(key, "level", StringComparison.Ordinal))) {
                    diagnostics.Add(Warning("ADF_HEADING_PROPERTIES_DROPPED", path, "ADF heading properties other than level cannot be represented in Markdown and were omitted."));
                }
                ReportMultilineHeadingText(node, path, diagnostics);
                builder.Append(new string('#', level)).Append(' ').Append(RenderInlineContent(node, path, diagnostics));
                break;
            case "codeBlock":
                string rawLanguage = node.GetStringAttribute("language") ?? string.Empty;
                string language = MarkdownFence.NormalizeLanguageToken(rawLanguage);
                if (node.ExtensionData.Count > 0 || node.Attributes.Keys.Any(key => !string.Equals(key, "language", StringComparison.Ordinal))) {
                    diagnostics.Add(Warning("ADF_CODE_PROPERTIES_DROPPED", path, "ADF code-block properties other than language cannot be represented in Markdown and were omitted."));
                }
                if (!string.Equals(language, rawLanguage, StringComparison.Ordinal)) {
                    diagnostics.Add(Warning("ADF_CODE_LANGUAGE_NORMALIZED", path, "ADF code-block language was normalized to one safe Markdown language token."));
                }
                string code = ExtractPlainText(node);
                if (code.IndexOf('\r') >= 0) {
                    diagnostics.Add(Warning("ADF_CODE_LINE_ENDINGS_NORMALIZED", path, "ADF code-block carriage-return line endings are normalized by Markdown round trips."));
                }
                string fence = MarkdownFence.BuildSafeFence(code);
                builder.Append(fence).Append(language).AppendLine();
                builder.Append(code);
                if (builder.Length > 0 && builder[builder.Length - 1] != '\n') builder.AppendLine();
                builder.Append(fence);
                break;
            case "blockquote":
                var quote = new StringBuilder();
                AppendChildrenAsBlocks(quote, node, path, options, diagnostics, listDepth);
                using (var reader = new StringReader(quote.ToString())) {
                    string? line;
                    bool first = true;
                    while ((line = reader.ReadLine()) != null) {
                        if (!first) builder.AppendLine();
                        builder.Append("> ").Append(line);
                        first = false;
                    }
                }
                break;
            case "bulletList":
            case "orderedList":
            case "taskList":
                AppendList(builder, node, path, options, diagnostics, listDepth);
                break;
            case "rule":
                builder.Append("---");
                break;
            case "blockCard":
                AppendSemanticFallback(builder, node, path, diagnostics);
                break;
            case "mediaSingle":
            case "mediaGroup":
                AppendMedia(builder, node, path, diagnostics);
                break;
            case "table":
                AppendTable(builder, node, path, diagnostics);
                break;
            case "panel":
                diagnostics.Add(Warning("ADF_PANEL_PROJECTED", path, "ADF panel styling is projected as a blockquote."));
                var panel = new StringBuilder();
                AppendChildrenAsBlocks(panel, node, path, options, diagnostics, listDepth);
                foreach (string line in panel.ToString().Split(new[] { "\r\n", "\n" }, StringSplitOptions.None)) {
                    if (builder.Length > 0 && builder[builder.Length - 1] != '\n') builder.AppendLine();
                    builder.Append("> ").Append(line);
                }
                break;
            default:
                diagnostics.Add(Warning("ADF_UNSUPPORTED_NODE", path, "ADF node '" + node.Type + "' is retained in the ADF model but has no exact Markdown projection."));
                if (node.Content.Count > 0) AppendChildrenAsBlocks(builder, node, path, options, diagnostics, listDepth);
                else if (!string.IsNullOrEmpty(ExtractPlainText(node))) builder.Append(MarkdownEscaper.EscapeLiteralText(ExtractPlainText(node)));
                else if (options.EmitUnsupportedPlaceholders) builder.Append("[Unsupported ADF node: ").Append(node.Type).Append(']');
                break;
        }
    }

    private static void AppendChildrenAsBlocks(StringBuilder builder, AdfNode node, string path, AdfConversionOptions options, List<AdfConversionDiagnostic> diagnostics, int listDepth) {
        for (int i = 0; i < node.Content.Count; i++) {
            if (i > 0) builder.AppendLine().AppendLine();
            AppendBlock(builder, node.Content[i], path + ".content[" + i + "]", options, diagnostics, listDepth);
        }
    }

    private static void AppendList(StringBuilder builder, AdfNode list, string path, AdfConversionOptions options, List<AdfConversionDiagnostic> diagnostics, int depth) {
        int order = list.GetInt32Attribute("order") ?? 1;
        if (list.Type == "taskList" &&
            (!string.IsNullOrWhiteSpace(list.GetStringAttribute("localId")) ||
             list.Content.Any(item => item.Type == "taskItem" && !string.IsNullOrWhiteSpace(item.GetStringAttribute("localId"))))) {
            diagnostics.Add(Warning(
                "ADF_TASK_LOCAL_IDS_REGENERATED",
                path,
                "Markdown has no task identity syntax; taskList and taskItem localId attributes were omitted and will be regenerated when Markdown is imported."));
        }
        for (int i = 0; i < list.Content.Count; i++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            AdfNode item = list.Content[i];
            if (i > 0) builder.AppendLine();
            if (item.Type == "taskList") {
                AppendList(builder, item, path + ".content[" + i + "]", options, diagnostics, depth + 1);
                continue;
            }
            ReportBlockProperties(item, path + ".content[" + i + "]", diagnostics);
            string marker = list.Type == "orderedList" ? (order + i).ToString(System.Globalization.CultureInfo.InvariantCulture) + ". " : "- ";
            if (item.Type == "taskItem") {
                bool done = string.Equals(item.GetStringAttribute("state"), "DONE", StringComparison.OrdinalIgnoreCase);
                marker += done ? "[x] " : "[ ] ";
            }
            builder.Append(new string(' ', depth * 2)).Append(marker);
            var itemBody = new StringBuilder();
            if (item.Type == "taskItem") {
                itemBody.Append(RenderInlineContent(item, path + ".content[" + i + "]", diagnostics));
            } else {
                AppendChildrenAsBlocks(itemBody, item, path + ".content[" + i + "]", options, diagnostics, depth + 1);
            }
            string[] lines = itemBody.ToString().Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            if (lines.Length > 0) builder.Append(lines[0]);
            for (int lineIndex = 1; lineIndex < lines.Length; lineIndex++) {
                builder.AppendLine().Append(new string(' ', depth * 2 + marker.Length)).Append(lines[lineIndex]);
            }
        }
    }

    private static void AppendTable(StringBuilder builder, AdfNode table, string path, List<AdfConversionDiagnostic> diagnostics) {
        if (table.Attributes.Count > 0 || table.ExtensionData.Count > 0) {
            diagnostics.Add(Warning("ADF_TABLE_ATTRIBUTES_DROPPED", path, "ADF table attributes cannot be represented in Markdown and were omitted."));
        }
        var rows = table.Content.Where(node => node.Type == "tableRow").ToList();
        if (rows.Count == 0) return;
        int columns = rows.Max(row => row.Content.Count);
        if (columns == 0) return;

        bool hasHeader = rows[0].Content.Any(cell => cell.Type == "tableHeader");
        bool hasMixedFirstRow = hasHeader && rows[0].Content.Any(cell => cell.Type != "tableHeader");
        bool hasHeaderAfterFirstRow = rows.Skip(1).Any(row => row.Content.Any(cell => cell.Type == "tableHeader"));
        if (!hasHeader) diagnostics.Add(Warning("ADF_TABLE_HEADER_SYNTHESIZED", path, "Markdown requires a header row; the first ADF row was used as the header."));
        if (hasMixedFirstRow || hasHeaderAfterFirstRow) {
            diagnostics.Add(Warning("ADF_TABLE_HEADER_LAYOUT_NORMALIZED", path, "Markdown supports only one all-header first row; source table header cell types were normalized."));
        }
        AppendTableRow(builder, rows[0], columns, path + ".content[0]", diagnostics);
        builder.AppendLine();
        builder.Append('|');
        for (int column = 0; column < columns; column++) builder.Append(" --- |");
        for (int row = 1; row < rows.Count; row++) {
            builder.AppendLine();
            AppendTableRow(builder, rows[row], columns, path + ".content[" + row + "]", diagnostics);
        }
    }

    private static void AppendTableRow(StringBuilder builder, AdfNode row, int columns, string path, List<AdfConversionDiagnostic> diagnostics) {
        ReportBlockProperties(row, path, diagnostics);
        builder.Append('|');
        for (int column = 0; column < columns; column++) {
            string value = column < row.Content.Count ? RenderCell(row.Content[column], path + ".content[" + column + "]", diagnostics) : string.Empty;
            builder.Append(' ').Append(value.Replace("\r", " ").Replace("\n", "<br/>")).Append(" |");
        }
    }

    private static string RenderCell(AdfNode cell, string path, List<AdfConversionDiagnostic> diagnostics) {
        ReportBlockProperties(cell, path, diagnostics);
        if (cell.Attributes.Count > 0 || cell.ExtensionData.Count > 0) {
            diagnostics.Add(Warning("ADF_TABLE_CELL_ATTRIBUTES_DROPPED", path, "ADF table-cell attributes cannot be represented in Markdown and were omitted."));
        }
        if (cell.Content.Count != 1 || cell.Content[0].Type != "paragraph") {
            diagnostics.Add(Warning("ADF_TABLE_CELL_BLOCKS_FLATTENED", path, "Markdown table cells reconstruct as one paragraph; the source cell block structure was flattened."));
        }
        var builder = new StringBuilder();
        for (int i = 0; i < cell.Content.Count; i++) {
            if (i > 0) builder.Append("<br/>");
            ReportBlockProperties(cell.Content[i], path + ".content[" + i + "]", diagnostics);
            builder.Append(RenderInlineContent(cell.Content[i], path + ".content[" + i + "]", diagnostics));
        }
        return builder.ToString();
    }

    private static string RenderInlineContent(AdfNode node, string path, List<AdfConversionDiagnostic> diagnostics) {
        var builder = new StringBuilder();
        for (int i = 0; i < node.Content.Count; i++) {
            var fragment = new StringBuilder();
            AppendInline(fragment, node.Content[i], path + ".content[" + i + "]", diagnostics);
            // Adjacent delimiter runs are parsed as one run by CommonMark. A comment
            // separates their syntax without adding visible whitespace or changing marks.
            if (builder.Length > 0 && fragment.Length > 0 && builder[builder.Length - 1] == fragment[0] &&
                fragment[0] is '*' or '_' or '~' or '^' or '`') builder.Append("<!-- -->");
            builder.Append(fragment);
        }
        return builder.ToString();
    }

    private static void ReportMultilineHeadingText(AdfNode heading, string path, List<AdfConversionDiagnostic> diagnostics) {
        for (int i = 0; i < heading.Content.Count; i++) {
            AdfNode child = heading.Content[i];
            if (child.Type == "text" && child.Text != null && (child.Text.IndexOf('\r') >= 0 || child.Text.IndexOf('\n') >= 0)) {
                diagnostics.Add(Warning(
                    "ADF_HEADING_LINE_BREAK_NORMALIZED",
                    path + ".content[" + i + "]",
                    "Markdown headings cannot preserve line endings inside one ADF text node; the projected heading structure is lossy."));
            }
        }
    }

    private static void AppendInline(StringBuilder builder, AdfNode node, string path, List<AdfConversionDiagnostic> diagnostics) {
        if (node.Type == "hardBreak") {
            ReportInlineProperties(node, path, diagnostics);
            builder.Append("<br />");
            return;
        }
        if (node.Type != "text") {
            AppendSemanticFallback(builder, node, path, diagnostics);
            return;
        }

        ReportInlineProperties(node, path, diagnostics);

        string rawText = node.Text ?? string.Empty;
        bool hasCode = node.Marks.Any(mark => string.Equals(mark.Type, "code", StringComparison.Ordinal));
        if (hasCode && (rawText.IndexOf('\r') >= 0 || rawText.IndexOf('\n') >= 0)) {
            diagnostics.Add(Warning("ADF_CODE_MARK_LINE_BREAK_NORMALIZED", path, "Markdown code spans normalize line breaks to spaces; the original ADF code-mark text cannot be represented exactly."));
        }
        bool hasDelimiterMark = node.Marks.Any(mark =>
            string.Equals(mark.Type, "strong", StringComparison.Ordinal) ||
            string.Equals(mark.Type, "em", StringComparison.Ordinal) ||
            string.Equals(mark.Type, "strike", StringComparison.Ordinal) ||
            string.Equals(mark.Type, "subsup", StringComparison.Ordinal));
        int leadingWhitespace = 0;
        while (leadingWhitespace < rawText.Length && char.IsWhiteSpace(rawText[leadingWhitespace])) leadingWhitespace++;
        int trailingWhitespace = rawText.Length;
        while (trailingWhitespace > leadingWhitespace && char.IsWhiteSpace(rawText[trailingWhitespace - 1])) trailingWhitespace--;
        if (hasDelimiterMark && (leadingWhitespace > 0 || trailingWhitespace < rawText.Length)) {
            diagnostics.Add(Warning("ADF_MARK_BOUNDARY_WHITESPACE_NORMALIZED", path, "Markdown delimiters cannot preserve boundary whitespace inside an ADF strong, emphasis, strike, superscript, or subscript mark; the whitespace was moved outside the marked span."));
            builder.Append(MarkdownEscaper.EscapeLiteralText(rawText.Substring(0, leadingWhitespace)));
            if (trailingWhitespace > leadingWhitespace) {
                builder.Append(RenderMarkedText(rawText.Substring(leadingWhitespace, trailingWhitespace - leadingWhitespace), node.Marks, path, diagnostics));
            }
            builder.Append(MarkdownEscaper.EscapeLiteralText(rawText.Substring(trailingWhitespace)));
            return;
        }
        builder.Append(RenderMarkedText(rawText, node.Marks, path, diagnostics));
    }

    private static string RenderMarkedText(string rawText, IEnumerable<AdfMark> marks, string path, List<AdfConversionDiagnostic> diagnostics) {
        AdfMark[] markList = marks.ToArray();
        foreach (AdfMark mark in markList) ReportMarkProperties(mark, path, diagnostics);
        bool hasCode = markList.Any(mark => string.Equals(mark.Type, "code", StringComparison.Ordinal));
        string? scriptType = markList
            .FirstOrDefault(mark => string.Equals(mark.Type, "subsup", StringComparison.Ordinal))
            ?.GetStringAttribute("type");
        string value = hasCode
            ? MarkdownFence.BuildSafeCodeSpan(rawText)
            : string.Equals(scriptType, "sup", StringComparison.OrdinalIgnoreCase)
                ? MarkdownEscaper.EscapeLiteralSuperscriptText(rawText)
                : string.Equals(scriptType, "sub", StringComparison.OrdinalIgnoreCase)
                    ? MarkdownEscaper.EscapeLiteralSubscriptText(rawText)
                    : MarkdownEscaper.EscapeLiteralText(rawText);
        var prefixes = new List<string>();
        var suffixes = new List<string>();
        void Wrap(string prefix, string suffix) { prefixes.Add(prefix); suffixes.Add(suffix); }
        foreach (AdfMark mark in markList.Where(mark => !string.Equals(mark.Type, "code", StringComparison.Ordinal))) {
            switch (mark.Type) {
                case "strong": Wrap("**", "**"); break;
                case "em": Wrap("*", "*"); break;
                case "strike": Wrap("~~", "~~"); break;
                case "underline": Wrap("<u>", "</u>"); break;
                case "subsup":
                    string? script = mark.GetStringAttribute("type");
                    if (string.Equals(script, "sup", StringComparison.OrdinalIgnoreCase)) Wrap("^", "^");
                    else if (string.Equals(script, "sub", StringComparison.OrdinalIgnoreCase)) Wrap("~", "~");
                    else diagnostics.Add(Warning("ADF_SUBSUP_TYPE_UNSUPPORTED", path, "ADF subsup marks require attrs.type 'sup' or 'sub'; the invalid mark was flattened."));
                    break;
                case "link":
                    string? href = mark.GetStringAttribute("href");
                    string? title = mark.GetStringAttribute("title");
                    if (mark.Attributes.Keys.Any(key => !string.Equals(key, "href", StringComparison.Ordinal) && !string.Equals(key, "title", StringComparison.Ordinal)) || mark.ExtensionData.Count > 0) {
                        diagnostics.Add(Warning("ADF_LINK_ATTRIBUTES_DROPPED", path, "ADF link attributes other than href and title cannot be represented in Markdown."));
                    }
                    if (!string.IsNullOrWhiteSpace(href)) {
                        string renderedTitle = MarkdownEscaper.FormatOptionalTitle(title);
                        Wrap("[", "](" + MarkdownEscaper.EscapeLinkUrl(href) + renderedTitle + ")");
                    }
                    break;
                default:
                    diagnostics.Add(Warning("ADF_UNSUPPORTED_MARK", path, "ADF mark '" + mark.Type + "' was flattened."));
                    break;
            }
        }
        if (prefixes.Count == 0) return value;
        // Compose once instead of copying the entire text again for each mark.
        var result = new StringBuilder();
        for (int i = prefixes.Count - 1; i >= 0; i--) result.Append(prefixes[i]);
        result.Append(value);
        foreach (string suffix in suffixes) result.Append(suffix);
        return result.ToString();
    }

    private static string ExtractPlainText(AdfNode node) {
        if (node.Type == "text") return node.Text ?? string.Empty;
        string? label = node.GetStringAttribute("text") ?? node.GetStringAttribute("shortName") ?? node.GetStringAttribute("url") ?? node.GetStringAttribute("alt");
        if (!string.IsNullOrEmpty(label)) return label!;
        var builder = new StringBuilder();
        foreach (AdfNode child in node.Content) builder.Append(ExtractPlainText(child));
        return builder.ToString();
    }

    private static bool EndsWithBlankLine(StringBuilder builder) => builder.Length >= 2 && builder[builder.Length - 1] == '\n' && builder[builder.Length - 2] == '\n';
    private static AdfConversionDiagnostic Warning(string code, string path, string message) => new AdfConversionDiagnostic(code, path, message, AdfConversionSeverity.Warning);
}
