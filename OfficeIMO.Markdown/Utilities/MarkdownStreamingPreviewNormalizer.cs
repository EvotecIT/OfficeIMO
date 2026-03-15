using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OfficeIMO.Markdown;

/// <summary>
/// Conservative markdown sanitizers for partial streaming preview text.
/// </summary>
/// <example>
/// <code>
/// var preview = MarkdownStreamingPreviewNormalizer.NormalizeIntelligenceXTranscript(partialDeltaText);
/// </code>
/// Use this only for partial or still-streaming transcript text where the host wants a conservative preview-safe cleanup pass.
/// For full document ingestion, prefer <see cref="MarkdownTranscriptPreparation"/> or <see cref="MarkdownInputNormalizer"/>.
/// </example>
public static class MarkdownStreamingPreviewNormalizer {
    private static readonly Regex ZeroWidthWhitespaceRegex = new(
        @"[\u200B\u2060\uFEFF]",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex LineStartMissingSpaceBeforeBoldBulletRegex = new(
        @"(?m)^(?<indent>\s*)-(?=\*\*)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex LineStartUnicodeDashBulletRegex = new(
        @"(?m)^(?<indent>\s*)[‐‑‒–—−](?=(?:\s*\*\*|[A-Z]{2,}\d+\b|[\p{Lu}][\p{L}\p{N}]{1,}\b))",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex LineStartHostLabelBulletRegex = new(
        @"(?m)^(?<indent>\s*)-(?=[A-Z]{2,}\d+\b)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex LeadingWhitespaceInsideStrongOpenRegex = new(
        @"(?:(?<=^)|(?<=[\s(\[{>]))\*\*[ \t]+(?=[^\s*\r\n])",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex CompactSignalFlowLabelSpacingRegex = new(
        @"->\s*(?:\*\*[^*\r\n]{1,120}:\*\*(?=\S)|[\p{L}][^:\r\n]{0,120}:(?=[^\s/\\]))",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex StandaloneHostLabelBulletRegex = new(
        @"^\s*-(?:\s*\*\*)?\s*[A-Z]{2,}\d+(?:\s*\*\*)?\s*:?\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex StructuralMarkdownLineRegex = new(
        @"^(?:[-+*]\s+|\d+[.)]\s+|#{1,6}\s+|>\s?|```|~~~|\|)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    /// <summary>
    /// Applies the explicit IntelligenceX transcript streaming-preview contract.
    /// </summary>
    /// <remarks>
    /// This path is intentionally conservative: it only performs line-start repairs for partial deltas,
    /// but escalates to the full explicit transcript input-normalization preset for known signal-flow and
    /// malformed strong-span artifacts.
    /// </remarks>
    public static string NormalizeIntelligenceXTranscript(string? text) {
        var normalized = text ?? string.Empty;
        if (normalized.Length == 0) {
            return string.Empty;
        }

        return MarkdownFence.ApplyTransformOutsideFencedCodeBlocks(normalized, static segment => {
            var value = ZeroWidthWhitespaceRegex.Replace(segment, string.Empty);
            if (RequiresIntelligenceXTranscriptFullCleanup(value)) {
                return MarkdownInlineCode.ApplyTransformPreservingInlineCodeSpans(
                    value,
                    static protectedInlineCode => MarkdownInputNormalizer.Normalize(
                        protectedInlineCode,
                        MarkdownInputNormalizationPresets.CreateIntelligenceXTranscript()));
            }

            value = LineStartUnicodeDashBulletRegex.Replace(value, "${indent}-");
            value = LineStartMissingSpaceBeforeBoldBulletRegex.Replace(value, "${indent}- ");
            value = LineStartHostLabelBulletRegex.Replace(value, "${indent}- ");
            value = LeadingWhitespaceInsideStrongOpenRegex.Replace(value, "**");
            value = MergeSplitHostLabelBullets(value);
            return value;
        });
    }

    private static bool RequiresIntelligenceXTranscriptFullCleanup(string text) {
        if (string.IsNullOrEmpty(text)) {
            return false;
        }

        return text.IndexOf("****", StringComparison.Ordinal) >= 0
               || text.IndexOf("->**", StringComparison.Ordinal) >= 0
               || HasCompactSignalFlowLabelSpacing(text);
    }

    private static bool HasCompactSignalFlowLabelSpacing(string text) {
        return !string.IsNullOrEmpty(text)
               && text.IndexOf("->", StringComparison.Ordinal) >= 0
               && CompactSignalFlowLabelSpacingRegex.IsMatch(text);
    }

    private static string MergeSplitHostLabelBullets(string text) {
        if (string.IsNullOrEmpty(text) || text.IndexOf('\n') < 0) {
            return text ?? string.Empty;
        }

        var hasCrLf = text.IndexOf("\r\n", StringComparison.Ordinal) >= 0;
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        if (lines.Length < 2) {
            return text;
        }

        var merged = new List<string>(lines.Length);
        for (var i = 0; i < lines.Length; i++) {
            var current = lines[i] ?? string.Empty;
            if (i + 1 < lines.Length
                && StandaloneHostLabelBulletRegex.IsMatch(current)
                && ShouldAttachHostLabelContinuation(lines[i + 1])) {
                var next = (lines[i + 1] ?? string.Empty).TrimStart();
                merged.Add(current.TrimEnd() + " " + next);
                i++;
                continue;
            }

            merged.Add(current);
        }

        var rebuilt = string.Join("\n", merged);
        return hasCrLf ? rebuilt.Replace("\n", "\r\n") : rebuilt;
    }

    private static bool ShouldAttachHostLabelContinuation(string line) {
        if (string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        var trimmed = line.TrimStart();
        return !StructuralMarkdownLineRegex.IsMatch(trimmed);
    }
}
