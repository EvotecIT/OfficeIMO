using System.Collections.Generic;

namespace OfficeIMO.Word.Markdown;

/// <summary>
/// Capability probes for markdown-to-Word contracts that may influence host adaptation.
/// </summary>
public static class MarkdownToWordCapabilities {
    /// <summary>
    /// Detects whether the current converter preserves isolated <c>Label: value</c> lines as separate
    /// narrative paragraphs when <see cref="MarkdownToWordOptions.PreferNarrativeSingleLineDefinitions"/> is enabled.
    /// </summary>
    public static bool PreservesNarrativeSingleLineDefinitionsAsSeparateParagraphs() {
        try {
            const string sampleMarkdown = """
                Status: healthy
                Impact: none
                """;

            using var document = sampleMarkdown.LoadFromMarkdown(
                MarkdownToWordPresets.CreateIntelligenceXTranscript());

            var bodyParagraphs = new List<string>();
            foreach (var paragraph in document.Paragraphs) {
                var text = string.Concat(paragraph.GetRuns().Select(run => run.Text));
                if (string.IsNullOrWhiteSpace(text)) {
                    continue;
                }

                bodyParagraphs.Add(text);
            }

            return bodyParagraphs.Contains("Status: healthy", StringComparer.Ordinal)
                   && bodyParagraphs.Contains("Impact: none", StringComparer.Ordinal);
        } catch {
            return false;
        }
    }
}
