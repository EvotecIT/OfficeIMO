using System.Text;
using OfficeIMO.Markdown;

namespace OfficeIMO.Adf;

internal static partial class AdfToMarkdownConverter {
    private static void ReportBlockProperties(AdfNode node, string path, List<AdfConversionDiagnostic> diagnostics) {
        if (node.MarkItems.Count > 0) diagnostics.Add(Omission("ADF_NODE_MARKS_DROPPED", path + ".marks", "ADF node-level marks cannot be represented in Markdown and were omitted."));
        string[]? retained = node.Type switch {
            "paragraph" or "blockquote" or "rule" or "bulletList" or "listItem" or "tableRow" => Array.Empty<string>(),
            "orderedList" => new[] { "order" },
            "taskList" => new[] { "localId" },
            "taskItem" => new[] { "localId", "state" },
            _ => null
        };
        if (retained != null && (node.ExtensionItems.Count > 0 || node.AttributeItems.Keys.Any(key => !retained.Contains(key, StringComparer.Ordinal)))) {
            diagnostics.Add(Omission("ADF_NODE_PROPERTIES_DROPPED", path, "ADF node properties outside the projected content cannot be represented in Markdown and were omitted."));
        }
    }

    private static void ReportInlineProperties(AdfNode node, string path, List<AdfConversionDiagnostic> diagnostics) {
        if (node.ExtensionItems.Count > 0 || node.AttributeItems.Count > 0 || (node.Type == "hardBreak" && node.MarkItems.Count > 0)) {
            diagnostics.Add(Omission("ADF_INLINE_PROPERTIES_DROPPED", path, "ADF inline properties outside text and supported marks cannot be represented in Markdown and were omitted."));
        }
    }

    private static void ReportMarkProperties(AdfMark mark, string path, List<AdfConversionDiagnostic> diagnostics) {
        if (mark.Type == "link") return; // Link-specific accounting includes href and title.
        if (mark.ExtensionItems.Count > 0 || mark.AttributeItems.Keys.Any(key => mark.Type != "subsup" || key != "type")) {
            diagnostics.Add(Omission("ADF_MARK_PROPERTIES_DROPPED", path, "ADF mark properties outside supported styling cannot be represented in Markdown and were omitted."));
        }
    }

    private static void AppendSemanticFallback(StringBuilder builder, AdfNode node, string path, List<AdfConversionDiagnostic> diagnostics) {
        string text = ExtractPlainText(node);
        diagnostics.Add(Omission(text.Length == 0 ? "ADF_UNSUPPORTED_INLINE_OMITTED" : "ADF_SEMANTIC_NODE_PROJECTED", path,
            text.Length == 0
                ? "ADF node '" + node.Type + "' has no visible fallback and was omitted."
                : "ADF node '" + node.Type + "' retains its visible fallback, but vendor identity and properties were omitted."));
        string? url = node.GetStringAttribute("url");
        if ((node.Type is "inlineCard" or "blockCard") && !string.IsNullOrWhiteSpace(url)) {
            builder.Append('[').Append(MarkdownEscaper.EscapeLiteralText(text)).Append("](").Append(MarkdownEscaper.EscapeLinkUrl(url!)).Append(')');
        } else {
            builder.Append(MarkdownEscaper.EscapeLiteralText(text));
        }
    }

    private static void AppendMedia(StringBuilder builder, AdfNode node, string path, List<AdfConversionDiagnostic> diagnostics) {
        diagnostics.Add(Omission("ADF_MEDIA_PROPERTIES_DROPPED", path, "Markdown media preserves external image URL and alternate text; ADF layout, identity and media properties were omitted."));
        for (int i = 0; i < node.ContentItems.Count; i++) {
            AdfNode media = node.ContentItems[i];
            if (i > 0) builder.AppendLine().AppendLine();
            string? url = media.GetStringAttribute("url");
            if (media.Type == "media" && media.GetStringAttribute("type") == "external" && !string.IsNullOrWhiteSpace(url)) {
                builder.Append("![").Append(MarkdownEscaper.EscapeLiteralText(media.GetStringAttribute("alt") ?? string.Empty))
                    .Append("](").Append(MarkdownEscaper.EscapeLinkUrl(url!)).Append(')');
                if (media.MarkItems.Count > 0) diagnostics.Add(Omission("ADF_MEDIA_MARKS_DROPPED", path + ".content[" + i + "].marks", "ADF media marks cannot be represented by this image projection and were omitted."));
            } else {
                AppendSemanticFallback(builder, media, path + ".content[" + i + "]", diagnostics);
            }
        }
    }

    private static AdfConversionDiagnostic Omission(string code, string path, string message) =>
        new AdfConversionDiagnostic(code, path, message, AdfConversionSeverity.Warning, OfficeConversionLossKind.Omission);
}
