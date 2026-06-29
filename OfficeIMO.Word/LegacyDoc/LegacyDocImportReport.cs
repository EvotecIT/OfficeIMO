using System.Text;
using OfficeIMO.Word.LegacyDoc.Diagnostics;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc {
    /// <summary>
    /// Compact import summary intended for corpus baselines and preflight checks.
    /// </summary>
    public sealed class LegacyDocImportReport {
        internal LegacyDocImportReport(LegacyDocDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            ParagraphCount = document.Paragraphs.Count;
            CharacterCount = document.Text.Length;
            DocumentPropertyCount = document.DocumentProperties.Count;
            DiagnosticCount = document.Diagnostics.Count;
            ErrorCount = document.Diagnostics.Count(diagnostic => diagnostic.Severity == LegacyDocDiagnosticSeverity.Error);
            WarningCount = document.Diagnostics.Count(diagnostic => diagnostic.Severity == LegacyDocDiagnosticSeverity.Warning);
            DiagnosticsByCode = document.Diagnostics
                .GroupBy(diagnostic => diagnostic.Code)
                .OrderBy(group => group.Key, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        }

        /// <summary>Gets the number of projected body paragraphs.</summary>
        public int ParagraphCount { get; }

        /// <summary>Gets the number of body text characters decoded from the piece table.</summary>
        public int CharacterCount { get; }

        /// <summary>Gets the number of projected built-in, application, and custom document properties.</summary>
        public int DocumentPropertyCount { get; }

        /// <summary>Gets the total number of diagnostics.</summary>
        public int DiagnosticCount { get; }

        /// <summary>Gets the number of error diagnostics.</summary>
        public int ErrorCount { get; }

        /// <summary>Gets the number of warning diagnostics.</summary>
        public int WarningCount { get; }

        /// <summary>Gets diagnostics grouped by code.</summary>
        public IReadOnlyDictionary<string, int> DiagnosticsByCode { get; }

        /// <summary>
        /// Formats the report as compact Markdown.
        /// </summary>
        public string ToMarkdown() {
            var builder = new StringBuilder();
            builder.AppendLine("# Legacy DOC Import Report");
            builder.AppendLine();
            builder.AppendLine("| Metric | Value |");
            builder.AppendLine("| --- | ---: |");
            builder.AppendLine($"| Paragraphs | {ParagraphCount} |");
            builder.AppendLine($"| Characters | {CharacterCount} |");
            builder.AppendLine($"| Document properties | {DocumentPropertyCount} |");
            builder.AppendLine($"| Diagnostics | {DiagnosticCount} |");
            builder.AppendLine($"| Errors | {ErrorCount} |");
            builder.AppendLine($"| Warnings | {WarningCount} |");

            if (DiagnosticsByCode.Count > 0) {
                builder.AppendLine();
                builder.AppendLine("## Diagnostics");
                builder.AppendLine();
                builder.AppendLine("| Code | Count |");
                builder.AppendLine("| --- | ---: |");
                foreach (KeyValuePair<string, int> entry in DiagnosticsByCode) {
                    builder.AppendLine($"| {entry.Key} | {entry.Value} |");
                }
            }

            return builder.ToString();
        }
    }
}
