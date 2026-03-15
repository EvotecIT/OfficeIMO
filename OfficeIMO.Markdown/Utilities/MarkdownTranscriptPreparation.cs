namespace OfficeIMO.Markdown;

using System;

/// <summary>
/// Explicit transcript-preparation helpers for hosts ingesting or exporting IntelligenceX-style markdown.
/// </summary>
public static class MarkdownTranscriptPreparation {
    /// <summary>
    /// Applies the shared IX transcript body preparation contract.
    /// </summary>
    /// <param name="markdown">Transcript markdown source.</param>
    /// <returns>Normalized transcript body markdown.</returns>
    public static string PrepareIntelligenceXTranscriptBody(string? markdown) {
        var value = markdown ?? string.Empty;
        if (value.Length == 0) {
            return string.Empty;
        }

        var normalized = MarkdownInputNormalizer.Normalize(
            value,
            MarkdownInputNormalizationPresets.CreateIntelligenceXTranscript());

        var prepared = string.IsNullOrEmpty(normalized) ? value : normalized;
        return MarkdownOrderedLists.SeparateAdjacentOrderedListItems(prepared);
    }

    /// <summary>
    /// Applies the shared IX transcript export-preparation contract after any host-specific transport markers
    /// have already been stripped.
    /// </summary>
    /// <param name="markdown">Transcript markdown source after host-specific marker cleanup.</param>
    /// <returns>Export-prepared transcript markdown.</returns>
    public static string PrepareIntelligenceXTranscriptForExport(string? markdown) {
        return MarkdownBlankLines.CollapseDuplicateBlankLines(
            PrepareIntelligenceXTranscriptBody(markdown));
    }

    /// <summary>
    /// Applies the shared IX transcript DOCX-preparation contract.
    /// </summary>
    /// <param name="markdown">Transcript markdown source.</param>
    /// <param name="preservesGroupedDefinitionLikeParagraphs">Whether the host renderer preserves grouped definition-like lines without compatibility repair.</param>
    /// <returns>DOCX-prepared transcript markdown.</returns>
    public static string PrepareIntelligenceXTranscriptForDocx(
        string? markdown,
        bool preservesGroupedDefinitionLikeParagraphs) {
        var prepared = PrepareIntelligenceXTranscriptBody(markdown);
        if (preservesGroupedDefinitionLikeParagraphs) {
            return prepared;
        }

        return MarkdownDefinitionLines.SeparateAdjacentDefinitionLikeLinesOutsideFencedCodeBlocks(prepared);
    }
}
