using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Legacy markdown migration helpers for older IntelligenceX transcript artifacts.
/// This owns compatibility cleanup only; alias fence registration stays in <see cref="MarkdownRendererIntelligenceXAdapter"/>.
/// </summary>
public static class MarkdownRendererIntelligenceXLegacyMigration {
    /// <summary>
    /// Registers the legacy IX transcript migration preprocessors if they are not already present.
    /// </summary>
    public static void Apply(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddPreProcessorIfMissing(options, NormalizeLegacyIxVisualArtifacts);
        AddPreProcessorIfMissing(options, NormalizeLegacyToolHeadingArtifacts);
    }

    /// <summary>
    /// Returns <see langword="true"/> when any IX legacy migration preprocessor is present.
    /// </summary>
    public static bool IsApplied(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        var processors = options.MarkdownPreProcessors;
        for (var i = 0; i < processors.Count; i++) {
            var processor = processors[i];
            if (processor == NormalizeLegacyIxVisualArtifacts || processor == NormalizeLegacyToolHeadingArtifacts) {
                return true;
            }
        }

        return false;
    }

    private static void AddPreProcessorIfMissing(MarkdownRendererOptions options, MarkdownTextPreProcessor processor) {
        var processors = options.MarkdownPreProcessors;
        for (int i = 0; i < processors.Count; i++) {
            if (processors[i] == processor) {
                return;
            }
        }

        processors.Add(processor);
    }

    private static readonly Regex LegacyToolHeadingBulletRegex = new(
        @"^(?<indent>\s*)-\s+(?<tool>[a-z0-9_.-]+):\s*(?<heading>#{2,6}\s+[^\r\n]+)\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Multiline);

    private static readonly Regex LegacyToolHeadingLeadRegex = new(
        @"^(?<indent>\s*)-\s+(?<tool>[a-z0-9_.-]+):\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Multiline);

    private static readonly Regex LegacyToolHeadingSplitBulletRegex = new(
        @"^(?<indent>\s*)-\s+(?<tool>[a-z0-9_.-]+):\s*(?<fragment>#{1,5})\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Multiline);

    private static readonly Regex LegacyToolSlugHeadingRegex = new(
        @"^(?<indent>\s*)#{2,6}\s+(?<tool>[a-z0-9_.-]+)\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Multiline);

    private static readonly Regex CachedToolEvidenceMarkerLineRegex = new(
        @"(?m)^[ \t]*ix:cached-tool-evidence:v1[ \t]*(?:\r?\n)?",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static string NormalizeLegacyIxVisualArtifacts(string markdown, MarkdownRendererOptions _) {
        if (string.IsNullOrEmpty(markdown)) {
            return markdown ?? string.Empty;
        }

        var value = StripCachedToolEvidenceMarkers(markdown);
        value = UpgradeLegacyVisualFences(value);
        value = UpgradeLegacyIndentedVisualBlocks(value);

        return value;
    }

    private static string StripCachedToolEvidenceMarkers(string markdown) {
        if (string.IsNullOrEmpty(markdown)
            || markdown.IndexOf("ix:cached-tool-evidence:v1", StringComparison.OrdinalIgnoreCase) < 0) {
            return markdown ?? string.Empty;
        }

        return CachedToolEvidenceMarkerLineRegex.Replace(markdown, string.Empty);
    }

    private static string NormalizeLegacyToolHeadingArtifacts(string markdown, MarkdownRendererOptions _) {
        if (string.IsNullOrEmpty(markdown)) {
            return markdown ?? string.Empty;
        }

        var hasCrLf = markdown.IndexOf("\r\n", StringComparison.Ordinal) >= 0;
        var normalized = markdown.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        var rewritten = new List<string>(lines.Length);
        var changed = false;

        for (var i = 0; i < lines.Length; i++) {
            var current = lines[i] ?? string.Empty;
            var bulletMatch = LegacyToolHeadingBulletRegex.Match(current);
            if (bulletMatch.Success) {
                rewritten.Add(bulletMatch.Groups["heading"].Value.Trim());
                changed = true;
                continue;
            }

            var bulletLeadMatch = LegacyToolHeadingLeadRegex.Match(current);
            if (bulletLeadMatch.Success && TryFindNextNonEmptyLine(lines, i + 1, out var promotedHeadingIndex)) {
                var next = lines[promotedHeadingIndex] ?? string.Empty;
                if (IsMarkdownHeadingLine(next)) {
                    changed = true;
                    continue;
                }
            }

            var splitBulletMatch = LegacyToolHeadingSplitBulletRegex.Match(current);
            if (splitBulletMatch.Success && TryFindNextNonEmptyLine(lines, i + 1, out var splitHeadingIndex)) {
                var next = lines[splitHeadingIndex] ?? string.Empty;
                if (TryParseMarkdownHeadingLine(next, out var nextDepth, out var headingText)) {
                    var fragmentDepth = splitBulletMatch.Groups["fragment"].Value.Length;
                    var combinedDepth = Math.Min(6, fragmentDepth + nextDepth);
                    rewritten.Add(new string('#', combinedDepth) + " " + headingText);
                    changed = true;
                    i = splitHeadingIndex;
                    continue;
                }
            }

            var slugMatch = LegacyToolSlugHeadingRegex.Match(current);
            if (slugMatch.Success && TryFindNextNonEmptyLine(lines, i + 1, out var nextIndex)) {
                var next = lines[nextIndex] ?? string.Empty;
                if (IsMarkdownHeadingLine(next)) {
                    changed = true;
                    continue;
                }
            }

            rewritten.Add(current);
        }

        if (!changed) {
            return markdown;
        }

        var rebuilt = string.Join("\n", rewritten);
        return hasCrLf ? rebuilt.Replace("\n", "\r\n") : rebuilt;
    }

    private static string UpgradeLegacyVisualFences(string input) {
        if (string.IsNullOrEmpty(input) || input.IndexOf("```", StringComparison.Ordinal) < 0) {
            return input ?? string.Empty;
        }

        var output = new StringBuilder(input.Length);
        var index = 0;
        while (index < input.Length) {
            var lineStart = index;
            while (index < input.Length && input[index] != '\r' && input[index] != '\n') {
                index++;
            }

            var lineEnd = index;
            if (index < input.Length && input[index] == '\r') {
                index++;
                if (index < input.Length && input[index] == '\n') {
                    index++;
                }
            } else if (index < input.Length && input[index] == '\n') {
                index++;
            }

            var line = input.Substring(lineStart, lineEnd - lineStart);
            var lineWithNewline = input.Substring(lineStart, index - lineStart);
            if (!MarkdownFence.TryReadContainerAwareFenceRun(line, out _, out var runMarker, out var runLength, out _)) {
                output.Append(lineWithNewline);
                continue;
            }

            var blockBuilder = new StringBuilder(lineWithNewline.Length + 256);
            blockBuilder.Append(lineWithNewline);
            var contentBuilder = new StringBuilder();
            var foundClosingFence = false;
            while (index < input.Length) {
                var innerLineStart = index;
                while (index < input.Length && input[index] != '\r' && input[index] != '\n') {
                    index++;
                }

                var innerLineEnd = index;
                if (index < input.Length && input[index] == '\r') {
                    index++;
                    if (index < input.Length && input[index] == '\n') {
                        index++;
                    }
                } else if (index < input.Length && input[index] == '\n') {
                    index++;
                }

                var innerLine = input.Substring(innerLineStart, innerLineEnd - innerLineStart);
                var innerLineWithNewline = input.Substring(innerLineStart, index - innerLineStart);
                if (MarkdownFence.TryReadContainerAwareFenceRun(innerLine, out _, out var closingMarker, out var closingLength, out var closingSuffix)
                    && closingMarker == runMarker
                    && closingLength >= runLength
                    && string.IsNullOrWhiteSpace(closingSuffix)) {
                    foundClosingFence = true;
                    blockBuilder.Append(innerLineWithNewline);
                    break;
                }

                contentBuilder.Append(innerLineWithNewline);
                blockBuilder.Append(innerLineWithNewline);
            }

            if (!foundClosingFence || !TryDetectLegacyIxVisualLanguageFromFence(line, contentBuilder.ToString(), out var targetLanguage)) {
                output.Append(blockBuilder);
                continue;
            }

            output.Append(RewriteFenceOpeningLine(lineWithNewline, targetLanguage));
            output.Append(contentBuilder);
            var originalBlock = blockBuilder.ToString();
            var openingLength = lineWithNewline.Length;
            var contentLength = contentBuilder.Length;
            output.Append(originalBlock.Substring(openingLength + contentLength));
        }

        return output.ToString();
    }

    private static string UpgradeLegacyIndentedVisualBlocks(string input) {
        if (string.IsNullOrEmpty(input)
            || input.IndexOf('{') < 0) {
            return input ?? string.Empty;
        }

        return ApplyTransformOutsideFencedCodeBlocks(input, static segment => {
            var hasCrLf = segment.IndexOf("\r\n", StringComparison.Ordinal) >= 0;
            var normalized = segment.Replace("\r\n", "\n").Replace('\r', '\n');
            var lines = normalized.Split('\n');
            var rewritten = new List<string>(lines.Length);
            var changed = false;

            for (var i = 0; i < lines.Length; i++) {
                var current = lines[i] ?? string.Empty;
                if (!IsLegacyIndentedCodeLine(current)) {
                    rewritten.Add(current);
                    continue;
                }

                var blockLines = new List<string>();
                var rawLines = new List<string>();
                var cursor = i;
                while (cursor < lines.Length) {
                    var line = lines[cursor] ?? string.Empty;
                    if (!IsLegacyIndentedCodeLine(line) && !string.IsNullOrWhiteSpace(line)) {
                        break;
                    }

                    rawLines.Add(line);
                    if (string.IsNullOrWhiteSpace(line)) {
                        blockLines.Add(string.Empty);
                    } else {
                        blockLines.Add(RemoveLegacyIndentedCodePrefix(line));
                    }

                    cursor++;
                }

                if (blockLines.Count == 0) {
                    rewritten.Add(current);
                    continue;
                }

                var candidate = string.Join("\n", blockLines);
                if (!TryDetectLegacyIxVisualLanguageFromJson(candidate, out var targetLanguage)) {
                    rewritten.AddRange(rawLines);
                    i = cursor - 1;
                    continue;
                }

                rewritten.Add("```" + targetLanguage);
                rewritten.AddRange(blockLines);
                rewritten.Add("```");
                changed = true;
                i = cursor - 1;
            }

            if (!changed) {
                return segment;
            }

            var rebuilt = string.Join("\n", rewritten);
            return hasCrLf ? rebuilt.Replace("\n", "\r\n") : rebuilt;
        });
    }

    private static bool TryDetectLegacyIxVisualLanguageFromFence(string openingFenceLine, string fenceContent, out string targetLanguage) {
        targetLanguage = string.Empty;
        if (string.IsNullOrWhiteSpace(fenceContent)) {
            return false;
        }

        if (!MarkdownFence.TryReadContainerAwareFenceRun(openingFenceLine, out _, out _, out _, out var suffix)) {
            return false;
        }

        var language = suffix.Trim();
        if (language.Length > 0 && !language.Equals("json", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        return TryDetectLegacyIxVisualLanguageFromJson(fenceContent, out targetLanguage);
    }

    private static bool TryDetectLegacyIxVisualLanguageFromJson(string candidate, out string targetLanguage) {
        targetLanguage = string.Empty;
        if (!MarkdownJsonVisualPayloadDetector.TryDetectSemanticKind(candidate, out var semanticKind)) {
            return false;
        }

        targetLanguage = semanticKind switch {
            MarkdownSemanticKinds.Chart => "ix-chart",
            MarkdownSemanticKinds.Network => "ix-network",
            MarkdownSemanticKinds.DataView => "ix-dataview",
            _ => string.Empty
        };
        return targetLanguage.Length > 0;
    }

    private static bool IsMarkdownHeadingLine(string line) {
        var trimmed = (line ?? string.Empty).TrimStart();
        if (trimmed.Length < 4 || trimmed[0] != '#') {
            return false;
        }

        var depth = 0;
        while (depth < trimmed.Length && trimmed[depth] == '#') {
            depth++;
        }

        return depth >= 2
               && depth <= 6
               && depth < trimmed.Length
               && char.IsWhiteSpace(trimmed[depth]);
    }

    private static bool TryParseMarkdownHeadingLine(string line, out int depth, out string text) {
        text = string.Empty;
        var trimmed = (line ?? string.Empty).TrimStart();
        if (trimmed.Length < 4 || trimmed[0] != '#') {
            depth = 0;
            return false;
        }

        depth = 0;
        while (depth < trimmed.Length && trimmed[depth] == '#') {
            depth++;
        }

        if (depth < 2 || depth > 6 || depth >= trimmed.Length || !char.IsWhiteSpace(trimmed[depth])) {
            depth = 0;
            return false;
        }

        text = trimmed.Substring(depth).TrimStart();
        return text.Length > 0;
    }

    private static bool TryFindNextNonEmptyLine(string[] lines, int startIndex, out int index) {
        for (var i = startIndex; i < lines.Length; i++) {
            if (!string.IsNullOrWhiteSpace(lines[i])) {
                index = i;
                return true;
            }
        }

        index = -1;
        return false;
    }

    private static bool IsLegacyIndentedCodeLine(string line) {
        if (string.IsNullOrEmpty(line)) {
            return false;
        }

        return line.StartsWith("    ", StringComparison.Ordinal)
               || line.StartsWith("\t", StringComparison.Ordinal);
    }

    private static string RemoveLegacyIndentedCodePrefix(string line) {
        if (line.StartsWith("\t", StringComparison.Ordinal)) {
            return line.Substring(1);
        }

        return line.StartsWith("    ", StringComparison.Ordinal) ? line.Substring(4) : line;
    }

    private static string RewriteFenceOpeningLine(string originalLineWithNewline, string language) {
        var newlineStart = originalLineWithNewline.Length;
        while (newlineStart > 0) {
            var ch = originalLineWithNewline[newlineStart - 1];
            if (ch != '\r' && ch != '\n') {
                break;
            }

            newlineStart--;
        }

        var line = originalLineWithNewline.Substring(0, newlineStart);
        var newline = originalLineWithNewline.Substring(newlineStart);
        if (!MarkdownFence.TryReadContainerAwareFenceRun(line, out var prefix, out var marker, out var runLength, out _)) {
            return originalLineWithNewline;
        }

        return prefix + new string(marker, runLength) + language + newline;
    }

    private static string ApplyTransformOutsideFencedCodeBlocks(string input, Func<string, string> transform) {
        if (string.IsNullOrEmpty(input)) {
            return input ?? string.Empty;
        }

        var output = new StringBuilder(input.Length);
        var outsideSegment = new StringBuilder();
        var inFence = false;
        char fenceMarker = '\0';
        var fenceRunLength = 0;

        var index = 0;
        while (index < input.Length) {
            var lineStart = index;
            while (index < input.Length && input[index] != '\r' && input[index] != '\n') {
                index++;
            }

            var lineEnd = index;
            if (index < input.Length && input[index] == '\r') {
                index++;
                if (index < input.Length && input[index] == '\n') {
                    index++;
                }
            } else if (index < input.Length && input[index] == '\n') {
                index++;
            }

            var line = input.Substring(lineStart, lineEnd - lineStart);
            var lineWithNewline = input.Substring(lineStart, index - lineStart);

            if (MarkdownFence.TryReadContainerAwareFenceRun(line, out _, out var runMarker, out var runLength, out var runSuffix)) {
                if (!inFence) {
                    FlushOutsideSegment(output, outsideSegment, transform);
                    inFence = true;
                    fenceMarker = runMarker;
                    fenceRunLength = runLength;
                    output.Append(lineWithNewline);
                    continue;
                }

                if (runMarker == fenceMarker && runLength >= fenceRunLength && string.IsNullOrWhiteSpace(runSuffix)) {
                    inFence = false;
                    fenceMarker = '\0';
                    fenceRunLength = 0;
                    output.Append(lineWithNewline);
                    continue;
                }
            }

            if (inFence) {
                output.Append(lineWithNewline);
            } else {
                outsideSegment.Append(lineWithNewline);
            }
        }

        FlushOutsideSegment(output, outsideSegment, transform);
        return output.ToString();
    }

    private static void FlushOutsideSegment(StringBuilder output, StringBuilder outsideSegment, Func<string, string> transform) {
        if (outsideSegment.Length == 0) {
            return;
        }

        output.Append(transform(outsideSegment.ToString()));
        outsideSegment.Clear();
    }
}
