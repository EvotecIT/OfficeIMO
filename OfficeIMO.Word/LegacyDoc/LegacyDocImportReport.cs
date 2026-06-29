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
            UnsupportedFeatureCount = document.UnsupportedFeatures.Count;
            DiagnosticCount = document.Diagnostics.Count;
            ErrorCount = document.Diagnostics.Count(diagnostic => diagnostic.Severity == LegacyDocDiagnosticSeverity.Error);
            WarningCount = document.Diagnostics.Count(diagnostic => diagnostic.Severity == LegacyDocDiagnosticSeverity.Warning);
            UnsupportedFeatures = document.UnsupportedFeatures.ToArray();
            UnsupportedFeaturesByCode = document.UnsupportedFeatures
                .GroupBy(feature => feature.Code)
                .OrderBy(group => group.Key, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
            UnsupportedFeaturesByKind = document.UnsupportedFeatures
                .GroupBy(feature => feature.Kind)
                .OrderBy(group => group.Key)
                .ToDictionary(group => group.Key, group => group.Count());
            UnsupportedFeaturesByDetail = document.UnsupportedFeatures
                .Select(GetUnsupportedFeatureDetailKey)
                .GroupBy(key => key, StringComparer.Ordinal)
                .OrderBy(group => group.Key, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
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

        /// <summary>Gets the number of unsupported or preserve-only features discovered while importing.</summary>
        public int UnsupportedFeatureCount { get; }

        /// <summary>Gets the total number of diagnostics.</summary>
        public int DiagnosticCount { get; }

        /// <summary>Gets the number of error diagnostics.</summary>
        public int ErrorCount { get; }

        /// <summary>Gets the number of warning diagnostics.</summary>
        public int WarningCount { get; }

        /// <summary>Gets diagnostics grouped by code.</summary>
        public IReadOnlyDictionary<string, int> DiagnosticsByCode { get; }

        /// <summary>Gets unsupported or preserve-only features discovered while importing.</summary>
        public IReadOnlyList<LegacyDocUnsupportedFeature> UnsupportedFeatures { get; }

        /// <summary>Gets unsupported features grouped by stable code.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedFeaturesByCode { get; }

        /// <summary>Gets unsupported features grouped by structured feature category.</summary>
        public IReadOnlyDictionary<LegacyDocUnsupportedFeatureKind, int> UnsupportedFeaturesByKind { get; }

        /// <summary>Gets unsupported features grouped by stable kind, code, and detail key.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedFeaturesByDetail { get; }

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
            builder.AppendLine($"| Unsupported features | {UnsupportedFeatureCount} |");
            builder.AppendLine($"| Diagnostics | {DiagnosticCount} |");
            builder.AppendLine($"| Errors | {ErrorCount} |");
            builder.AppendLine($"| Warnings | {WarningCount} |");

            if (UnsupportedFeatures.Count > 0) {
                builder.AppendLine();
                builder.AppendLine("## Unsupported Features");
                builder.AppendLine();
                builder.AppendLine("| Kind | Code | Detail | Entry |");
                builder.AppendLine("| --- | --- | --- | --- |");
                foreach (LegacyDocUnsupportedFeature feature in UnsupportedFeatures) {
                    builder.AppendLine($"| {feature.Kind} | {feature.Code} | {feature.DetailCode ?? string.Empty} | {feature.EntryPath ?? string.Empty} |");
                }
            }

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

        private static string GetUnsupportedFeatureDetailKey(LegacyDocUnsupportedFeature feature) {
            return string.Join("|", new[] {
                feature.Kind.ToString(),
                feature.Code,
                feature.DetailCode ?? string.Empty
            });
        }
    }
}
