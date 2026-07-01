namespace OfficeIMO.Word {
    /// <summary>
    /// Writes deterministic report formats for <see cref="WordComparisonResult"/> without coupling callers to the comparer.
    /// </summary>
    public static class WordComparisonReportWriter {
        /// <summary>
        /// Serializes a comparison result to deterministic JSON.
        /// </summary>
        /// <param name="result">Comparison result to serialize.</param>
        /// <returns>JSON report text with stable property order.</returns>
        public static string ToJson(WordComparisonResult result) {
            if (result == null) {
                throw new ArgumentNullException(nameof(result));
            }

            var builder = new StringBuilder();
            builder.AppendLine("{");
            AppendJsonProperty(builder, 1, "sourcePath", result.SourcePath, comma: true);
            AppendJsonProperty(builder, 1, "targetPath", result.TargetPath, comma: true);
            AppendJsonProperty(builder, 1, "hasChanges", result.HasChanges, comma: true);
            AppendJsonProperty(builder, 1, "findingCount", result.Findings.Count, comma: true);
            AppendJsonSummary(builder, result, comma: true);
            AppendJsonFindings(builder, result, comma: false);
            builder.Append('}');
            return builder.ToString();
        }

        /// <summary>
        /// Renders a comparison result as Markdown suitable for review notes and automation logs.
        /// </summary>
        /// <param name="result">Comparison result to render.</param>
        /// <returns>Markdown report text.</returns>
        public static string ToMarkdown(WordComparisonResult result) {
            if (result == null) {
                throw new ArgumentNullException(nameof(result));
            }

            var builder = new StringBuilder();
            builder.AppendLine("# Word Comparison Report");
            builder.AppendLine();
            builder.AppendLine("## Summary");
            builder.AppendLine();
            builder.AppendLine("| Metric | Value |");
            builder.AppendLine("| --- | --- |");
            builder.Append("| Source | ");
            builder.Append(EscapeMarkdownCell(result.SourcePath));
            builder.AppendLine(" |");
            builder.Append("| Target | ");
            builder.Append(EscapeMarkdownCell(result.TargetPath));
            builder.AppendLine(" |");
            builder.Append("| Has changes | ");
            builder.Append(result.HasChanges ? "yes" : "no");
            builder.AppendLine(" |");
            builder.Append("| Findings | ");
            builder.Append(result.Findings.Count.ToString(System.Globalization.CultureInfo.InvariantCulture));
            builder.AppendLine(" |");
            builder.AppendLine();

            AppendSummaryMarkdown(builder, "By Scope", result.Findings.GroupBy(finding => finding.Scope.ToString()).OrderBy(group => group.Key, StringComparer.Ordinal));
            AppendSummaryMarkdown(builder, "By Change", result.Findings.GroupBy(finding => finding.ChangeKind.ToString()).OrderBy(group => group.Key, StringComparer.Ordinal));
            AppendFindingsMarkdown(builder, result);

            return builder.ToString().TrimEnd();
        }

        /// <summary>
        /// Returns a compact single-line comparison summary suitable for CLI wrappers and CI annotations.
        /// </summary>
        /// <param name="result">Comparison result to summarize.</param>
        /// <returns>Single-line summary text.</returns>
        public static string ToTextSummary(WordComparisonResult result) {
            if (result == null) {
                throw new ArgumentNullException(nameof(result));
            }

            if (result.Findings.Count == 0) {
                return "No structural differences detected.";
            }

            string count = result.Findings.Count.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string noun = result.Findings.Count == 1 ? "finding" : "findings";
            string scopeSummary = string.Join(
                ", ",
                result.Findings
                    .GroupBy(finding => finding.Scope)
                    .OrderBy(group => group.Key)
                    .Select(group => group.Key + "=" + group.Count().ToString(System.Globalization.CultureInfo.InvariantCulture)));
            string changeSummary = string.Join(
                ", ",
                result.Findings
                    .GroupBy(finding => finding.ChangeKind)
                    .OrderBy(group => group.Key)
                    .Select(group => group.Key + "=" + group.Count().ToString(System.Globalization.CultureInfo.InvariantCulture)));

            return count + " " + noun + " detected. Scopes: " + scopeSummary + ". Changes: " + changeSummary + ".";
        }

        private static void AppendJsonSummary(StringBuilder builder, WordComparisonResult result, bool comma) {
            AppendIndent(builder, 1);
            AppendJsonString(builder, "summary");
            builder.AppendLine(": {");
            AppendJsonCountObject(builder, 2, "byScope", result.Findings.GroupBy(finding => finding.Scope.ToString()).OrderBy(group => group.Key, StringComparer.Ordinal), comma: true);
            AppendJsonCountObject(builder, 2, "byChange", result.Findings.GroupBy(finding => finding.ChangeKind.ToString()).OrderBy(group => group.Key, StringComparer.Ordinal), comma: false);
            AppendIndent(builder, 1);
            builder.Append('}');
            builder.AppendLine(comma ? "," : string.Empty);
        }

        private static void AppendJsonFindings(StringBuilder builder, WordComparisonResult result, bool comma) {
            AppendIndent(builder, 1);
            AppendJsonString(builder, "findings");
            builder.AppendLine(": [");
            for (int i = 0; i < result.Findings.Count; i++) {
                WordComparisonFinding finding = result.Findings[i];
                AppendIndent(builder, 2);
                builder.AppendLine("{");
                AppendJsonProperty(builder, 3, "scope", finding.Scope.ToString(), comma: true);
                AppendJsonProperty(builder, 3, "changeKind", finding.ChangeKind.ToString(), comma: true);
                AppendJsonProperty(builder, 3, "location", finding.Location, comma: true);
                AppendJsonProperty(builder, 3, "detailedLocation", finding.DetailedLocation, comma: true);
                AppendJsonProperty(builder, 3, "sourceIndex", finding.SourceIndex, comma: true);
                AppendJsonProperty(builder, 3, "targetIndex", finding.TargetIndex, comma: true);
                AppendJsonProperty(builder, 3, "sourceText", finding.SourceText, comma: true);
                AppendJsonProperty(builder, 3, "targetText", finding.TargetText, comma: true);
                AppendJsonProperty(builder, 3, "message", finding.Message, comma: false);
                AppendIndent(builder, 2);
                builder.Append('}');
                builder.AppendLine(i == result.Findings.Count - 1 ? string.Empty : ",");
            }

            AppendIndent(builder, 1);
            builder.Append(']');
            builder.AppendLine(comma ? "," : string.Empty);
        }

        private static void AppendJsonCountObject(
            StringBuilder builder,
            int depth,
            string name,
            IEnumerable<IGrouping<string, WordComparisonFinding>> groups,
            bool comma) {
            AppendIndent(builder, depth);
            AppendJsonString(builder, name);
            builder.AppendLine(": {");
            List<IGrouping<string, WordComparisonFinding>> groupList = groups.ToList();
            for (int i = 0; i < groupList.Count; i++) {
                IGrouping<string, WordComparisonFinding> group = groupList[i];
                AppendJsonProperty(builder, depth + 1, group.Key, group.Count(), comma: i < groupList.Count - 1);
            }

            AppendIndent(builder, depth);
            builder.Append('}');
            builder.AppendLine(comma ? "," : string.Empty);
        }

        private static void AppendSummaryMarkdown(
            StringBuilder builder,
            string title,
            IEnumerable<IGrouping<string, WordComparisonFinding>> groups) {
            builder.Append("## ");
            builder.AppendLine(title);
            builder.AppendLine();

            List<IGrouping<string, WordComparisonFinding>> groupList = groups.ToList();
            if (groupList.Count == 0) {
                builder.AppendLine("_None._");
                builder.AppendLine();
                return;
            }

            builder.AppendLine("| Name | Count |");
            builder.AppendLine("| --- | ---: |");
            foreach (IGrouping<string, WordComparisonFinding> group in groupList) {
                builder.Append("| ");
                builder.Append(EscapeMarkdownCell(group.Key));
                builder.Append(" | ");
                builder.Append(group.Count().ToString(System.Globalization.CultureInfo.InvariantCulture));
                builder.AppendLine(" |");
            }

            builder.AppendLine();
        }

        private static void AppendFindingsMarkdown(StringBuilder builder, WordComparisonResult result) {
            builder.AppendLine("## Findings");
            builder.AppendLine();
            if (result.Findings.Count == 0) {
                builder.AppendLine("_None._");
                return;
            }

            builder.AppendLine("| # | Scope | Change | Location | Detailed Location | Source | Target | Message |");
            builder.AppendLine("| ---: | --- | --- | --- | --- | --- | --- | --- |");
            for (int i = 0; i < result.Findings.Count; i++) {
                WordComparisonFinding finding = result.Findings[i];
                builder.Append("| ");
                builder.Append(i.ToString(System.Globalization.CultureInfo.InvariantCulture));
                builder.Append(" | ");
                builder.Append(finding.Scope);
                builder.Append(" | ");
                builder.Append(finding.ChangeKind);
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(finding.Location));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(finding.DetailedLocation));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(finding.SourceText));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(finding.TargetText));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(finding.Message));
                builder.AppendLine(" |");
            }
        }

        private static void AppendJsonProperty(StringBuilder builder, int depth, string name, string? value, bool comma) {
            AppendIndent(builder, depth);
            AppendJsonString(builder, name);
            builder.Append(": ");
            if (value == null) {
                builder.Append("null");
            } else {
                AppendJsonString(builder, value);
            }

            builder.AppendLine(comma ? "," : string.Empty);
        }

        private static void AppendJsonProperty(StringBuilder builder, int depth, string name, int value, bool comma) {
            AppendIndent(builder, depth);
            AppendJsonString(builder, name);
            builder.Append(": ");
            builder.Append(value.ToString(System.Globalization.CultureInfo.InvariantCulture));
            builder.AppendLine(comma ? "," : string.Empty);
        }

        private static void AppendJsonProperty(StringBuilder builder, int depth, string name, int? value, bool comma) {
            AppendIndent(builder, depth);
            AppendJsonString(builder, name);
            builder.Append(": ");
            if (value.HasValue) {
                builder.Append(value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
            } else {
                builder.Append("null");
            }

            builder.AppendLine(comma ? "," : string.Empty);
        }

        private static void AppendJsonProperty(StringBuilder builder, int depth, string name, bool value, bool comma) {
            AppendIndent(builder, depth);
            AppendJsonString(builder, name);
            builder.Append(": ");
            builder.Append(value ? "true" : "false");
            builder.AppendLine(comma ? "," : string.Empty);
        }

        private static void AppendJsonString(StringBuilder builder, string value) {
            builder.Append('"');
            foreach (char ch in value) {
                switch (ch) {
                    case '"':
                        builder.Append("\\\"");
                        break;
                    case '\\':
                        builder.Append("\\\\");
                        break;
                    case '\b':
                        builder.Append("\\b");
                        break;
                    case '\f':
                        builder.Append("\\f");
                        break;
                    case '\n':
                        builder.Append("\\n");
                        break;
                    case '\r':
                        builder.Append("\\r");
                        break;
                    case '\t':
                        builder.Append("\\t");
                        break;
                    default:
                        if (char.IsControl(ch)) {
                            builder.Append("\\u");
                            builder.Append(((int)ch).ToString("x4", System.Globalization.CultureInfo.InvariantCulture));
                        } else {
                            builder.Append(ch);
                        }

                        break;
                }
            }

            builder.Append('"');
        }

        private static void AppendIndent(StringBuilder builder, int depth) {
            builder.Append(' ', depth * 2);
        }

        private static string EscapeMarkdownCell(string? value) =>
            EscapeMarkdownText(value ?? string.Empty)
                .Replace("|", "\\|");

        private static string EscapeMarkdownText(string value) =>
            value.Replace("\r", " ").Replace("\n", " ");
    }
}
