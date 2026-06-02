using OfficeIMO.Markdown;

namespace OfficeIMO.Markup;

public static partial class OfficeMarkupParser {
    private static MarkdownReaderOptions CreateNestedMarkdownOptions() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        RegisterOfficeFences(options);
        return options;
    }

    private static string ToPlainText(InlineSequence sequence) {
        var sb = new StringBuilder();
        ((IPlainTextMarkdownInline)sequence).AppendPlainText(sb);
        return sb.ToString();
    }

    private static bool IsMermaid(string language) =>
        string.Equals(language, "mermaid", StringComparison.OrdinalIgnoreCase);

    private static string NormalizeCommand(string value) =>
        (value ?? string.Empty).Trim().Replace("_", "-").ToLowerInvariant();

    private static void CopyAttributes(IDictionary<string, string> source, IDictionary<string, string> target) {
        foreach (var pair in source) {
            target[pair.Key] = pair.Value;
        }
    }

    private static void ApplyPlacement(OfficeMarkupBlock block, IDictionary<string, string> attributes) {
        var placement = CreatePlacement(attributes);
        if (placement == null) {
            return;
        }

        switch (block) {
            case OfficeMarkupImageBlock image:
                image.Placement = placement;
                break;
            case OfficeMarkupDiagramBlock diagram:
                diagram.Placement = placement;
                break;
            case OfficeMarkupChartBlock chart:
                chart.Placement = placement;
                break;
            case OfficeMarkupTextBoxBlock textBox:
                textBox.Placement = placement;
                break;
            case OfficeMarkupColumnsBlock columns:
                columns.Placement = placement;
                break;
            case OfficeMarkupCardBlock card:
                card.Placement = placement;
                break;
        }
    }

    private static OfficeMarkupPlacement? CreatePlacement(IDictionary<string, string> attributes) {
        var placement = new OfficeMarkupPlacement {
            X = GetAttribute(attributes, "x"),
            Y = GetAttribute(attributes, "y"),
            Width = GetAttribute(attributes, "w") ?? GetAttribute(attributes, "width"),
            Height = GetAttribute(attributes, "h") ?? GetAttribute(attributes, "height")
        };

        return placement.HasValue ? placement : null;
    }

    private static string? GetAttribute(OfficeMarkupDirective directive, string name) {
        return directive.Attributes.TryGetValue(name, out var value) ? value : null;
    }

    private static bool TryGetInt32(OfficeMarkupDirective directive, string name, out int value) {
        value = 0;
        var text = GetAttribute(directive, name);
        return !string.IsNullOrWhiteSpace(text) && int.TryParse(text, out value);
    }

    private static bool TryGetBoolean(OfficeMarkupDirective directive, string name, out bool value) {
        value = false;
        var text = GetAttribute(directive, name);
        if (string.IsNullOrWhiteSpace(text)) {
            return false;
        }

        if (bool.TryParse(text, out value)) {
            return true;
        }

        if (string.Equals(text, "yes", StringComparison.OrdinalIgnoreCase) || string.Equals(text, "1", StringComparison.OrdinalIgnoreCase)) {
            value = true;
            return true;
        }

        if (string.Equals(text, "no", StringComparison.OrdinalIgnoreCase) || string.Equals(text, "0", StringComparison.OrdinalIgnoreCase)) {
            value = false;
            return true;
        }

        return false;
    }

    private static IEnumerable<IReadOnlyList<string>> ParseDelimitedRows(string body) {
        if (string.IsNullOrWhiteSpace(body)) {
            yield break;
        }

        var lines = body.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        foreach (var line in lines) {
            if (string.IsNullOrWhiteSpace(line)) {
                continue;
            }

            var separator = line.IndexOf('\t') >= 0 ? '\t' : ',';
            yield return line.Split(separator).Select(cell => cell.Trim()).ToArray();
        }
    }
}
