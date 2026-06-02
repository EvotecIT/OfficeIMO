using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownInputNormalizer {
    private static string NormalizeHostLabelBulletArtifacts(string text) {
        if (string.IsNullOrEmpty(text)) {
            return text ?? string.Empty;
        }

        var spaced = LineStartUnicodeDashBulletRegex.Replace(text, "${indent}-");
        spaced = LineStartMissingSpaceBeforeBoldBulletRegex.Replace(spaced, "${indent}- ");
        spaced = LineStartBoldBulletStrongOpenWhitespaceRegex.Replace(spaced, "${lead}");
        spaced = LineStartHostLabelBulletRegex.Replace(spaced, "${indent}- ");
        if (spaced.IndexOf('\n') < 0 && spaced.IndexOf('\r') < 0) {
            return spaced;
        }

        var hasCrLf = spaced.IndexOf("\r\n", StringComparison.Ordinal) >= 0;
        var normalized = spaced.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        if (lines.Length < 2) {
            return spaced;
        }

        var merged = new List<string>(lines.Length);
        var changed = !spaced.Equals(text, StringComparison.Ordinal);
        for (var i = 0; i < lines.Length; i++) {
            var current = lines[i] ?? string.Empty;
            if (i + 1 < lines.Length
                && StandaloneHostLabelBulletRegex.IsMatch(current)
                && ShouldAttachHostLabelContinuation(lines[i + 1])) {
                var next = (lines[i + 1] ?? string.Empty).TrimStart();
                merged.Add(current.TrimEnd() + " " + next);
                changed = true;
                i++;
                continue;
            }

            merged.Add(current);
        }

        if (!changed) {
            return text;
        }

        var rebuilt = string.Join("\n", merged);
        return hasCrLf ? rebuilt.Replace("\n", "\r\n") : rebuilt;
    }

    private static string RemoveStandaloneHashSeparatorsBeforeHeadings(string text) {
        if (string.IsNullOrEmpty(text) || text.IndexOf('#') < 0) {
            return text ?? string.Empty;
        }

        var hasCrLf = text.IndexOf("\r\n", StringComparison.Ordinal) >= 0;
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        var rewritten = new List<string>(lines.Length);
        var changed = false;

        for (var i = 0; i < lines.Length; i++) {
            var current = lines[i] ?? string.Empty;
            if (StandaloneSingleHashSeparatorRegex.IsMatch(current)
                && TryFindNextNonEmptyLine(lines, i + 1, out var nextIndex)
                && IsMarkdownHeadingLine(lines[nextIndex] ?? string.Empty)) {
                changed = true;
                i = nextIndex - 1;
                continue;
            }

            rewritten.Add(current);
        }

        if (!changed) {
            return text;
        }

        var rebuilt = string.Join("\n", rewritten);
        return hasCrLf ? rebuilt.Replace("\n", "\r\n") : rebuilt;
    }

    private static bool IsMarkdownHeadingLine(string line) {
        var trimmed = line.TrimStart();
        if (trimmed.Length < 4 || trimmed[0] != '#') {
            return false;
        }

        var depth = 0;
        while (depth < trimmed.Length && trimmed[depth] == '#') {
            depth++;
        }

        return depth is >= 2 and <= 6
               && depth < trimmed.Length
               && char.IsWhiteSpace(trimmed[depth]);
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

    private static bool ShouldAttachHostLabelContinuation(string line) {
        if (string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        var trimmed = line.TrimStart();
        return !StructuralMarkdownLineRegex.IsMatch(trimmed);
    }
}
