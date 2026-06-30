using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Word {
    public sealed partial class WordContentControlFormValidationResult {
        /// <summary>
        /// Serializes this validation result to deterministic JSON for CI, service logs, and thin wrapper surfaces.
        /// </summary>
        public string ToJson() {
            var builder = new StringBuilder();
            builder.AppendLine("{");
            AppendJsonProperty(builder, 1, "isValid", IsValid, comma: true);
            AppendJsonProperty(builder, 1, "expectedKeyCount", ExpectedKeys.Count, comma: true);
            AppendJsonProperty(builder, 1, "suppliedKeyCount", SuppliedKeys.Count, comma: true);
            AppendJsonProperty(builder, 1, "issueCount", Issues.Count, comma: true);
            AppendJsonStringArray(builder, 1, "expectedKeys", ExpectedKeys, comma: true);
            AppendJsonStringArray(builder, 1, "suppliedKeys", SuppliedKeys, comma: true);
            AppendJsonIssues(builder, 1, "issues", Issues, comma: false);
            builder.Append('}');
            return builder.ToString();
        }

        /// <summary>
        /// Renders this validation result as Markdown suitable for review notes and automation artifacts.
        /// </summary>
        public string ToMarkdown() {
            var builder = new StringBuilder();
            builder.AppendLine("# Content-Control Form Validation");
            builder.AppendLine();
            builder.AppendLine("## Summary");
            builder.AppendLine();
            builder.AppendLine("| Metric | Value |");
            builder.AppendLine("| --- | ---: |");
            builder.AppendLine("| Valid | " + (IsValid ? "yes" : "no") + " |");
            builder.AppendLine("| Expected keys | " + ExpectedKeys.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) + " |");
            builder.AppendLine("| Supplied keys | " + SuppliedKeys.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) + " |");
            builder.AppendLine("| Issues | " + Issues.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) + " |");
            builder.AppendLine();

            AppendNameList(builder, "Expected Keys", ExpectedKeys);
            AppendNameList(builder, "Supplied Keys", SuppliedKeys);
            AppendIssueMarkdown(builder);

            return builder.ToString().TrimEnd();
        }

        private static void AppendNameList(StringBuilder builder, string title, IReadOnlyList<string> names) {
            builder.AppendLine("## " + title);
            builder.AppendLine();
            if (names.Count == 0) {
                builder.AppendLine("_None._");
                builder.AppendLine();
                return;
            }

            foreach (string name in names) {
                builder.Append("- ");
                builder.AppendLine(EscapeMarkdown(name));
            }

            builder.AppendLine();
        }

        private void AppendIssueMarkdown(StringBuilder builder) {
            builder.AppendLine("## Issues");
            builder.AppendLine();
            if (Issues.Count == 0) {
                builder.AppendLine("_None._");
                builder.AppendLine();
                return;
            }

            builder.AppendLine("| Kind | Key | Control Type | Message |");
            builder.AppendLine("| --- | --- | --- | --- |");
            foreach (WordContentControlFormIssue issue in Issues) {
                builder.Append("| ");
                builder.Append(EscapeMarkdown(issue.Kind.ToString()));
                builder.Append(" | ");
                builder.Append(EscapeMarkdown(issue.Key ?? string.Empty));
                builder.Append(" | ");
                builder.Append(EscapeMarkdown(issue.ControlType));
                builder.Append(" | ");
                builder.Append(EscapeMarkdown(issue.Message));
                builder.AppendLine(" |");
            }

            builder.AppendLine();
        }

        private static void AppendJsonIssues(StringBuilder builder, int indent, string name, IReadOnlyList<WordContentControlFormIssue> issues, bool comma) {
            AppendIndent(builder, indent);
            builder.Append('"');
            builder.Append(name);
            builder.AppendLine("\": [");
            for (int i = 0; i < issues.Count; i++) {
                WordContentControlFormIssue issue = issues[i];
                AppendIndent(builder, indent + 1);
                builder.Append("{ ");
                AppendInlineJsonProperty(builder, "kind", issue.Kind.ToString(), comma: true);
                builder.Append(' ');
                AppendInlineJsonProperty(builder, "key", issue.Key, comma: true);
                builder.Append(' ');
                AppendInlineJsonProperty(builder, "controlType", issue.ControlType, comma: true);
                builder.Append(' ');
                AppendInlineJsonProperty(builder, "message", issue.Message, comma: false);
                builder.Append(" }");
                if (i < issues.Count - 1) {
                    builder.Append(',');
                }

                builder.AppendLine();
            }

            AppendIndent(builder, indent);
            builder.Append(']');
            if (comma) {
                builder.Append(',');
            }

            builder.AppendLine();
        }

        private static void AppendJsonStringArray(StringBuilder builder, int indent, string name, IReadOnlyList<string> values, bool comma) {
            AppendIndent(builder, indent);
            builder.Append('"');
            builder.Append(name);
            builder.Append("\": [");
            for (int i = 0; i < values.Count; i++) {
                if (i > 0) {
                    builder.Append(", ");
                }

                AppendJsonString(builder, values[i]);
            }

            builder.Append(']');
            if (comma) {
                builder.Append(',');
            }

            builder.AppendLine();
        }

        private static void AppendJsonProperty(StringBuilder builder, int indent, string name, int value, bool comma) {
            AppendIndent(builder, indent);
            builder.Append('"');
            builder.Append(name);
            builder.Append("\": ");
            builder.Append(value.ToString(System.Globalization.CultureInfo.InvariantCulture));
            if (comma) {
                builder.Append(',');
            }

            builder.AppendLine();
        }

        private static void AppendJsonProperty(StringBuilder builder, int indent, string name, bool value, bool comma) {
            AppendIndent(builder, indent);
            builder.Append('"');
            builder.Append(name);
            builder.Append("\": ");
            builder.Append(value ? "true" : "false");
            if (comma) {
                builder.Append(',');
            }

            builder.AppendLine();
        }

        private static void AppendInlineJsonProperty(StringBuilder builder, string name, string? value, bool comma) {
            builder.Append('"');
            builder.Append(name);
            builder.Append("\": ");
            if (value == null) {
                builder.Append("null");
            } else {
                AppendJsonString(builder, value);
            }

            if (comma) {
                builder.Append(',');
            }
        }

        private static void AppendJsonString(StringBuilder builder, string value) {
            builder.Append('"');
            foreach (char character in value) {
                switch (character) {
                    case '\\':
                        builder.Append("\\\\");
                        break;
                    case '"':
                        builder.Append("\\\"");
                        break;
                    case '\r':
                        builder.Append("\\r");
                        break;
                    case '\n':
                        builder.Append("\\n");
                        break;
                    case '\t':
                        builder.Append("\\t");
                        break;
                    default:
                        builder.Append(character);
                        break;
                }
            }

            builder.Append('"');
        }

        private static string EscapeMarkdown(string value) {
            return value.Replace("|", "\\|");
        }

        private static void AppendIndent(StringBuilder builder, int indent) {
            builder.Append(' ', indent * 2);
        }
    }
}
