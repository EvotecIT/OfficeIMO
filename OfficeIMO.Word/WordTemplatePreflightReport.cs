using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Word {
    /// <summary>
    /// Describes template binding capabilities checked by <see cref="WordTemplatePreflightReport"/>.
    /// </summary>
    public enum WordTemplatePreflightCapability {
        /// <summary>All template binding operations requested by the supplied data can run without known template issues.</summary>
        BindTemplate,

        /// <summary>MERGEFIELD values can be bound without missing supplied values.</summary>
        BindMergeFields,

        /// <summary>Conditional block markers can be evaluated without missing values or marker structure issues.</summary>
        BindConditionalBlocks,

        /// <summary>Repeated block markers can be expanded without missing rows or marker structure issues.</summary>
        BindRepeatingBlocks
    }

    /// <summary>
    /// Summarizes whether a Word mail-merge template can be safely bound with supplied data.
    /// </summary>
    public sealed class WordTemplatePreflightReport {
        private readonly WordMailMergeTemplateInspection _inspection;

        private WordTemplatePreflightReport(WordMailMergeTemplateInspection inspection) {
            _inspection = inspection ?? throw new ArgumentNullException(nameof(inspection));
            MergeFieldNames = inspection.MergeFieldNames.ToArray();
            ConditionalBlockNames = inspection.ConditionalBlockNames.ToArray();
            RepeatingBlockNames = inspection.RepeatingBlockNames.ToArray();
            Issues = inspection.Issues.ToArray();
        }

        /// <summary>Unique MERGEFIELD names found in the template.</summary>
        public IReadOnlyList<string> MergeFieldNames { get; }

        /// <summary>Unique conditional block names found in the template.</summary>
        public IReadOnlyList<string> ConditionalBlockNames { get; }

        /// <summary>Unique repeated block names found in the template.</summary>
        public IReadOnlyList<string> RepeatingBlockNames { get; }

        /// <summary>Validation issues found during template inspection.</summary>
        public IReadOnlyList<WordMailMergeTemplateIssue> Issues { get; }

        /// <summary>Number of unique MERGEFIELD names found in the template.</summary>
        public int MergeFieldCount => MergeFieldNames.Count;

        /// <summary>Number of unique conditional block names found in the template.</summary>
        public int ConditionalBlockCount => ConditionalBlockNames.Count;

        /// <summary>Number of unique repeated block names found in the template.</summary>
        public int RepeatingBlockCount => RepeatingBlockNames.Count;

        /// <summary>Number of validation issues found during template inspection.</summary>
        public int IssueCount => Issues.Count;

        /// <summary>True when no known issue prevents binding the template with the supplied data.</summary>
        public bool CanBindTemplate => Can(WordTemplatePreflightCapability.BindTemplate);

        /// <summary>
        /// Creates a preflight report from an existing template inspection.
        /// </summary>
        /// <param name="inspection">Template inspection returned by <see cref="WordMailMerge.InspectTemplate"/>.</param>
        public static WordTemplatePreflightReport FromInspection(WordMailMergeTemplateInspection inspection) {
            return new WordTemplatePreflightReport(inspection);
        }

        /// <summary>
        /// Returns whether the requested template binding capability has no blocking diagnostics.
        /// </summary>
        /// <param name="capability">Capability to check.</param>
        public bool Can(WordTemplatePreflightCapability capability) {
            return GetDiagnostics(capability).Count == 0;
        }

        /// <summary>
        /// Returns diagnostics that block the requested template binding capability.
        /// </summary>
        /// <param name="capability">Capability to inspect.</param>
        public IReadOnlyList<WordMailMergeTemplateIssue> GetDiagnostics(WordTemplatePreflightCapability capability) {
            switch (capability) {
                case WordTemplatePreflightCapability.BindTemplate:
                    return Issues.ToArray();
                case WordTemplatePreflightCapability.BindMergeFields:
                    return Issues.Where(issue => issue.Kind == WordMailMergeTemplateIssueKind.MissingMergeFieldValue).ToArray();
                case WordTemplatePreflightCapability.BindConditionalBlocks:
                    return Issues.Where(IsConditionalIssue).ToArray();
                case WordTemplatePreflightCapability.BindRepeatingBlocks:
                    return Issues.Where(IsRepeatingIssue).ToArray();
                default:
                    throw new ArgumentOutOfRangeException(nameof(capability), capability, "Unsupported template preflight capability.");
            }
        }

        /// <summary>
        /// Throws when the requested capability has blocking diagnostics, otherwise returns this report.
        /// </summary>
        /// <param name="capability">Capability to enforce.</param>
        public WordTemplatePreflightReport EnsureCan(WordTemplatePreflightCapability capability) {
            IReadOnlyList<WordMailMergeTemplateIssue> diagnostics = GetDiagnostics(capability);
            if (diagnostics.Count > 0) {
                throw new InvalidOperationException(string.Join(Environment.NewLine, diagnostics.Select(issue => issue.Message)));
            }

            return this;
        }

        /// <summary>
        /// Serializes this report to deterministic JSON for CI and service logs.
        /// </summary>
        public string ToJson() {
            var builder = new StringBuilder();
            builder.AppendLine("{");
            AppendJsonProperty(builder, 1, "canBindTemplate", CanBindTemplate, comma: true);
            AppendJsonProperty(builder, 1, "mergeFieldCount", MergeFieldCount, comma: true);
            AppendJsonProperty(builder, 1, "conditionalBlockCount", ConditionalBlockCount, comma: true);
            AppendJsonProperty(builder, 1, "repeatingBlockCount", RepeatingBlockCount, comma: true);
            AppendJsonProperty(builder, 1, "issueCount", IssueCount, comma: true);
            AppendJsonStringArray(builder, 1, "mergeFieldNames", MergeFieldNames, comma: true);
            AppendJsonStringArray(builder, 1, "conditionalBlockNames", ConditionalBlockNames, comma: true);
            AppendJsonStringArray(builder, 1, "repeatingBlockNames", RepeatingBlockNames, comma: true);
            AppendJsonIssues(builder, 1, "issues", Issues, comma: false);
            builder.Append('}');
            return builder.ToString();
        }

        /// <summary>
        /// Renders this report as Markdown suitable for review notes and automation logs.
        /// </summary>
        public string ToMarkdown() {
            var builder = new StringBuilder();
            builder.AppendLine("# Word Template Preflight Report");
            builder.AppendLine();
            builder.AppendLine("## Summary");
            builder.AppendLine();
            builder.AppendLine("| Metric | Value |");
            builder.AppendLine("| --- | ---: |");
            builder.AppendLine($"| Can bind template | {(CanBindTemplate ? "yes" : "no")} |");
            builder.AppendLine($"| Merge fields | {MergeFieldCount} |");
            builder.AppendLine($"| Conditional blocks | {ConditionalBlockCount} |");
            builder.AppendLine($"| Repeating blocks | {RepeatingBlockCount} |");
            builder.AppendLine($"| Issues | {IssueCount} |");
            builder.AppendLine();

            AppendNameList(builder, "Merge Fields", MergeFieldNames);
            AppendNameList(builder, "Conditional Blocks", ConditionalBlockNames);
            AppendNameList(builder, "Repeating Blocks", RepeatingBlockNames);
            AppendIssueMarkdown(builder);

            return builder.ToString().TrimEnd();
        }

        private static bool IsConditionalIssue(WordMailMergeTemplateIssue issue) {
            return issue.Kind == WordMailMergeTemplateIssueKind.MissingConditionalValue
                || issue.Kind == WordMailMergeTemplateIssueKind.UnmatchedConditionalStart
                || issue.Kind == WordMailMergeTemplateIssueKind.UnmatchedConditionalEnd
                || issue.Kind == WordMailMergeTemplateIssueKind.MismatchedConditionalEnd;
        }

        private static bool IsRepeatingIssue(WordMailMergeTemplateIssue issue) {
            return issue.Kind == WordMailMergeTemplateIssueKind.MissingRepeatingBlockData
                || issue.Kind == WordMailMergeTemplateIssueKind.UnmatchedRepeatingBlockStart
                || issue.Kind == WordMailMergeTemplateIssueKind.UnmatchedRepeatingBlockEnd
                || issue.Kind == WordMailMergeTemplateIssueKind.MismatchedRepeatingBlockEnd;
        }

        private static void AppendNameList(StringBuilder builder, string title, IReadOnlyList<string> names) {
            builder.AppendLine($"## {title}");
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

            builder.AppendLine("| Kind | Name | Message |");
            builder.AppendLine("| --- | --- | --- |");
            foreach (WordMailMergeTemplateIssue issue in Issues) {
                builder.Append("| ");
                builder.Append(EscapeMarkdown(issue.Kind.ToString()));
                builder.Append(" | ");
                builder.Append(EscapeMarkdown(issue.Name));
                builder.Append(" | ");
                builder.Append(EscapeMarkdown(issue.Message));
                builder.AppendLine(" |");
            }

            builder.AppendLine();
        }

        private static void AppendJsonIssues(StringBuilder builder, int indent, string name, IReadOnlyList<WordMailMergeTemplateIssue> issues, bool comma) {
            AppendIndent(builder, indent);
            builder.Append('"');
            builder.Append(name);
            builder.AppendLine("\": [");
            for (int i = 0; i < issues.Count; i++) {
                WordMailMergeTemplateIssue issue = issues[i];
                AppendIndent(builder, indent + 1);
                builder.Append("{ ");
                AppendInlineJsonProperty(builder, "kind", issue.Kind.ToString(), comma: true);
                builder.Append(' ');
                AppendInlineJsonProperty(builder, "name", issue.Name, comma: true);
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
            builder.Append(value);
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
            AppendJsonString(builder, value);
            if (comma) {
                builder.Append(',');
            }
        }

        private static void AppendJsonString(StringBuilder builder, string? value) {
            builder.Append('"');
            if (!string.IsNullOrEmpty(value)) {
                foreach (char character in value!) {
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
                            if (char.IsControl(character)) {
                                builder.Append("\\u");
                                builder.Append(((int)character).ToString("x4", CultureInfo.InvariantCulture));
                            } else {
                                builder.Append(character);
                            }

                            break;
                    }
                }
            }

            builder.Append('"');
        }

        private static void AppendIndent(StringBuilder builder, int indent) {
            builder.Append(' ', indent * 2);
        }

        private static string EscapeMarkdown(string? value) {
            return string.IsNullOrEmpty(value)
                ? string.Empty
                : value!.Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
        }
    }
}
