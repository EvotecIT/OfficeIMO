using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownInputNormalizer {
    private static string NormalizeCompactFenceBodyBoundaries(string input) {
        if (string.IsNullOrEmpty(input)) {
            return input ?? string.Empty;
        }

        var output = new StringBuilder(input.Length + 64);
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
            var newline = input.Substring(lineEnd, index - lineEnd);

            if (!inFence && TryNormalizeCompactFenceOpeningLine(line, out var normalizedLine, out var normalizedMarker, out var normalizedRunLength)) {
                inFence = true;
                fenceMarker = normalizedMarker;
                fenceRunLength = normalizedRunLength;
                output.Append(normalizedLine);
                output.Append(newline);
                continue;
            }

            if (MarkdownFence.TryReadContainerAwareFenceRun(line, out _, out var runMarker, out var runLength, out var runSuffix)) {
                if (!inFence) {
                    inFence = true;
                    fenceMarker = runMarker;
                    fenceRunLength = runLength;
                } else if (runMarker == fenceMarker && runLength >= fenceRunLength && string.IsNullOrWhiteSpace(runSuffix)) {
                    inFence = false;
                    fenceMarker = '\0';
                    fenceRunLength = 0;
                }
            }

            output.Append(line);
            output.Append(newline);
        }

        return output.ToString();
    }

    private static bool TryNormalizeCompactFenceOpeningLine(string line, out string normalizedLine, out char fenceMarker, out int fenceRunLength) {
        normalizedLine = line ?? string.Empty;
        fenceMarker = '\0';
        fenceRunLength = 0;

        if (string.IsNullOrEmpty(line)) {
            return false;
        }

        if (!MarkdownFence.TryReadContainerAwareFenceRun(line, out var linePrefix, out fenceMarker, out fenceRunLength, out var runSuffix)) {
            return false;
        }

        if (string.IsNullOrWhiteSpace(runSuffix)) {
            return false;
        }

        var suffix = runSuffix.TrimStart();
        if (suffix.Length == 0) {
            return false;
        }

        if (!TrySplitCompactFenceSuffix(suffix, out var language, out var body)) {
            return false;
        }

        normalizedLine = linePrefix + new string(fenceMarker, fenceRunLength) + language + "\n" + linePrefix + body;
        return true;
    }

    private static bool TrySplitCompactFenceSuffix(string suffix, out string language, out string body) {
        language = string.Empty;
        body = string.Empty;

        foreach (var candidate in CompactFenceLanguages) {
            if (!suffix.StartsWith(candidate, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            var remainder = suffix.Substring(candidate.Length);
            if (string.IsNullOrWhiteSpace(remainder)) {
                continue;
            }

            if (!LooksLikeCompactFenceBody(candidate, remainder)) {
                continue;
            }

            language = candidate;
            body = remainder;
            return true;
        }

        return false;
    }

    private static bool LooksLikeCompactFenceBody(string language, string remainder) {
        if (string.IsNullOrWhiteSpace(remainder)) {
            return false;
        }

        var trimmed = remainder.TrimStart();
        if (trimmed.Length == 0) {
            return false;
        }

        if (string.Equals(language, "mermaid", StringComparison.OrdinalIgnoreCase)) {
            foreach (var prefix in MermaidBodyPrefixes) {
                if (trimmed.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }
            }

            return false;
        }

        if (string.Equals(language, "json", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(language, "jsonc", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(language, "json5", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(language, "chart", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(language, "ix-chart", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(language, "network", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(language, "visnetwork", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(language, "ix-network", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(language, "ix-dataview", StringComparison.OrdinalIgnoreCase)) {
            return trimmed[0] == '{' || trimmed[0] == '[';
        }

        return false;
    }
}
