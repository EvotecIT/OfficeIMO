using OfficeIMO.Markdown;

namespace OfficeIMO.Markup;

public static partial class OfficeMarkupParser {
    private sealed class OfficeMarkupDirective {
        private OfficeMarkupDirective(string command, Dictionary<string, string> attributes, string body) {
            Command = command;
            Attributes = attributes;
            Body = body;
        }

        public string Command { get; }
        public Dictionary<string, string> Attributes { get; }
        public string Body { get; }

        public static OfficeMarkupDirective Parse(string language, string content) {
            var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var normalizedContent = (content ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
            var lines = normalizedContent.Split('\n');
            var command = CommandFromLanguage(language);
            var index = 0;

            while (index < lines.Length && string.IsNullOrWhiteSpace(lines[index])) {
                index++;
            }

            if (string.IsNullOrEmpty(command) && index < lines.Length) {
                var firstTokens = OfficeMarkupParser.Tokenize(lines[index]);
                if (firstTokens.Count > 0) {
                    command = firstTokens[0];
                    OfficeMarkupParser.AddInlineAttributes(firstTokens.Skip(1), attributes);
                    index++;
                }
            } else if (index < lines.Length) {
                var firstTokens = OfficeMarkupParser.Tokenize(lines[index]);
                if (firstTokens.Count > 0 && LooksLikeAttributes(firstTokens)) {
                    OfficeMarkupParser.AddInlineAttributes(firstTokens, attributes);
                    index++;
                }
            }

            while (index < lines.Length) {
                var line = lines[index];
                if (string.IsNullOrWhiteSpace(line)) {
                    index++;
                    break;
                }

                if (!OfficeMarkupParser.TryParseAttributeLine(line, attributes)) {
                    break;
                }

                index++;
            }

            var body = string.Join("\n", lines.Skip(index)).Trim('\n');
            return new OfficeMarkupDirective(command, attributes, body);
        }

        private static string CommandFromLanguage(string language) {
            var value = (language ?? string.Empty).Trim();
            if (!value.StartsWith("officeimo-", StringComparison.OrdinalIgnoreCase)) {
                return string.Empty;
            }

            var suffix = value.Substring("officeimo-".Length);
            if (string.Equals(suffix, "presentation", StringComparison.OrdinalIgnoreCase)
                || string.Equals(suffix, "document", StringComparison.OrdinalIgnoreCase)
                || string.Equals(suffix, "workbook", StringComparison.OrdinalIgnoreCase)) {
                return string.Empty;
            }

            return suffix;
        }

        private static bool LooksLikeAttributes(IReadOnlyList<string> tokens) {
            for (int i = 0; i < tokens.Count; i++) {
                if (tokens[i].IndexOf('=') < 0) {
                    return false;
                }
            }

            return tokens.Count > 0;
        }

    }
}
