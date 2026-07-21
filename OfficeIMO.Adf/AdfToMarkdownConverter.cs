using System.IO;
using System.Text;

namespace OfficeIMO.Adf;

internal static class AdfToMarkdownConverter {
    internal static string Convert(AdfDocument document, AdfConversionOptions options, List<AdfConversionDiagnostic> diagnostics) {
        var builder = new StringBuilder();
        for (int i = 0; i < document.Content.Count; i++) {
            AppendBlock(builder, document.Content[i], "$.content[" + i + "]", options, diagnostics, 0);
            if (i < document.Content.Count - 1 && !EndsWithBlankLine(builder)) builder.AppendLine().AppendLine();
        }
        return builder.ToString().TrimEnd();
    }

    private static void AppendBlock(StringBuilder builder, AdfNode node, string path, AdfConversionOptions options, List<AdfConversionDiagnostic> diagnostics, int listDepth) {
        switch (node.Type) {
            case "paragraph":
                builder.Append(RenderInlineContent(node, path, diagnostics));
                break;
            case "heading":
                int level = node.GetInt32Attribute("level") ?? 1;
                level = Math.Max(1, Math.Min(6, level));
                builder.Append(new string('#', level)).Append(' ').Append(RenderInlineContent(node, path, diagnostics));
                break;
            case "codeBlock":
                string language = node.GetStringAttribute("language") ?? string.Empty;
                builder.Append("```").Append(language).AppendLine();
                builder.Append(ExtractPlainText(node));
                if (builder.Length > 0 && builder[builder.Length - 1] != '\n') builder.AppendLine();
                builder.Append("```");
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
                else if (!string.IsNullOrEmpty(node.Text)) builder.Append(EscapeText(node.Text!));
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
        for (int i = 0; i < list.Content.Count; i++) {
            AdfNode item = list.Content[i];
            if (i > 0) builder.AppendLine();
            string marker = list.Type == "orderedList" ? (order + i).ToString(System.Globalization.CultureInfo.InvariantCulture) + ". " : "- ";
            if (item.Type == "taskItem") {
                bool done = item.Attributes.TryGetValue("state", out var state) && string.Equals(state.GetString(), "DONE", StringComparison.OrdinalIgnoreCase);
                marker += done ? "[x] " : "[ ] ";
            }
            builder.Append(new string(' ', depth * 2)).Append(marker);
            var itemBody = new StringBuilder();
            AppendChildrenAsBlocks(itemBody, item, path + ".content[" + i + "]", options, diagnostics, depth + 1);
            string[] lines = itemBody.ToString().Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            if (lines.Length > 0) builder.Append(lines[0]);
            for (int lineIndex = 1; lineIndex < lines.Length; lineIndex++) {
                builder.AppendLine().Append(new string(' ', depth * 2 + marker.Length)).Append(lines[lineIndex]);
            }
        }
    }

    private static void AppendTable(StringBuilder builder, AdfNode table, string path, List<AdfConversionDiagnostic> diagnostics) {
        var rows = table.Content.Where(node => node.Type == "tableRow").ToList();
        if (rows.Count == 0) return;
        int columns = rows.Max(row => row.Content.Count);
        if (columns == 0) return;

        bool hasHeader = rows[0].Content.Any(cell => cell.Type == "tableHeader");
        if (!hasHeader) diagnostics.Add(Warning("ADF_TABLE_HEADER_SYNTHESIZED", path, "Markdown requires a header row; the first ADF row was used as the header."));
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
        builder.Append('|');
        for (int column = 0; column < columns; column++) {
            string value = column < row.Content.Count ? RenderCell(row.Content[column], path + ".content[" + column + "]", diagnostics) : string.Empty;
            builder.Append(' ').Append(value.Replace("|", "\\|").Replace("\r", " ").Replace("\n", "<br/>")).Append(" |");
        }
    }

    private static string RenderCell(AdfNode cell, string path, List<AdfConversionDiagnostic> diagnostics) {
        var builder = new StringBuilder();
        for (int i = 0; i < cell.Content.Count; i++) {
            if (i > 0) builder.Append("<br/>");
            builder.Append(RenderInlineContent(cell.Content[i], path + ".content[" + i + "]", diagnostics));
        }
        return builder.ToString();
    }

    private static string RenderInlineContent(AdfNode node, string path, List<AdfConversionDiagnostic> diagnostics) {
        var builder = new StringBuilder();
        for (int i = 0; i < node.Content.Count; i++) AppendInline(builder, node.Content[i], path + ".content[" + i + "]", diagnostics);
        return builder.ToString();
    }

    private static void AppendInline(StringBuilder builder, AdfNode node, string path, List<AdfConversionDiagnostic> diagnostics) {
        if (node.Type == "hardBreak") {
            builder.Append("  \n");
            return;
        }
        if (node.Type != "text") {
            diagnostics.Add(Warning("ADF_UNSUPPORTED_INLINE", path, "ADF inline node '" + node.Type + "' was flattened to text."));
            builder.Append(EscapeText(ExtractPlainText(node)));
            return;
        }

        string value = EscapeText(node.Text ?? string.Empty);
        foreach (AdfMark mark in node.Marks) {
            switch (mark.Type) {
                case "strong": value = "**" + value + "**"; break;
                case "em": value = "*" + value + "*"; break;
                case "strike": value = "~~" + value + "~~"; break;
                case "code": value = "`" + (node.Text ?? string.Empty).Replace("`", "\\`") + "`"; break;
                case "link":
                    string? href = mark.GetStringAttribute("href");
                    value = string.IsNullOrWhiteSpace(href) ? value : "[" + value + "](" + href!.Replace(")", "\\)") + ")";
                    break;
                default:
                    diagnostics.Add(Warning("ADF_UNSUPPORTED_MARK", path, "ADF mark '" + mark.Type + "' was flattened."));
                    break;
            }
        }
        builder.Append(value);
    }

    private static string ExtractPlainText(AdfNode node) {
        if (node.Type == "text") return node.Text ?? string.Empty;
        var builder = new StringBuilder();
        foreach (AdfNode child in node.Content) builder.Append(ExtractPlainText(child));
        return builder.ToString();
    }

    private static string EscapeText(string value) => (value ?? string.Empty)
        .Replace("\\", "\\\\").Replace("*", "\\*").Replace("_", "\\_").Replace("[", "\\[").Replace("]", "\\]");

    private static bool EndsWithBlankLine(StringBuilder builder) => builder.Length >= 2 && builder[builder.Length - 1] == '\n' && builder[builder.Length - 2] == '\n';
    private static AdfConversionDiagnostic Warning(string code, string path, string message) => new AdfConversionDiagnostic(code, path, message, AdfConversionSeverity.Warning);
}
