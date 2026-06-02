using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownInputNormalizer {
    private static string NormalizeSignalFlowLabelSpacing(string text) {
        if (string.IsNullOrEmpty(text) || text.IndexOf("->", StringComparison.Ordinal) < 0) {
            return text ?? string.Empty;
        }

        var hasCrLf = text.IndexOf("\r\n", StringComparison.Ordinal) >= 0;
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        var changed = false;

        for (var i = 0; i < lines.Length; i++) {
            var line = lines[i] ?? string.Empty;
            if (line.IndexOf("->", StringComparison.Ordinal) < 0
                || line.IndexOf('`') >= 0) {
                continue;
            }

            var rewritten = NormalizeSignalFlowArrowSegments(line);
            if (rewritten.Equals(line, StringComparison.Ordinal)) {
                continue;
            }

            lines[i] = rewritten;
            changed = true;
        }

        if (!changed) {
            return text;
        }

        var rebuilt = string.Join("\n", lines);
        return hasCrLf ? rebuilt.Replace("\n", "\r\n") : rebuilt;
    }

    private static string NormalizeSignalFlowArrowSegments(string line) {
        var segments = line.Split(new[] { "->" }, StringSplitOptions.None);
        if (segments.Length < 2) {
            return line;
        }

        var builder = new StringBuilder(line.Length + 8);
        builder.Append(segments[0]);
        for (var i = 1; i < segments.Length; i++) {
            builder.Append("->");
            builder.Append(NormalizeSignalFlowSegmentLabelSpacing(segments[i]));
        }

        return builder.ToString();
    }

    private static string NormalizeSignalFlowSegmentLabelSpacing(string segment) {
        if (string.IsNullOrEmpty(segment)) {
            return segment ?? string.Empty;
        }

        var start = 0;
        while (start < segment.Length && char.IsWhiteSpace(segment[start])) {
            start++;
        }

        if (start >= segment.Length) {
            return segment;
        }

        var strongNormalized = TryNormalizeLeadingStrongSignalLabel(segment, start);
        if (!strongNormalized.Equals(segment, StringComparison.Ordinal)) {
            return strongNormalized;
        }

        return TryNormalizeLeadingPlainSignalLabel(segment, start);
    }

    private static string TryNormalizeLeadingStrongSignalLabel(string segment, int start) {
        if (start + 1 >= segment.Length || segment[start] != '*' || segment[start + 1] != '*') {
            return segment;
        }

        var close = segment.IndexOf("**", start + 2, segment.Length - (start + 2), StringComparison.Ordinal);
        if (close < 0 || close + 2 >= segment.Length) {
            return segment;
        }

        if (segment[close - 1] != ':') {
            return segment;
        }

        var next = segment[close + 2];
        if (char.IsWhiteSpace(next)) {
            return segment;
        }

        return segment.Insert(close + 2, " ");
    }

    private static string TryNormalizeLeadingPlainSignalLabel(string segment, int start) {
        var candidate = segment.Substring(start);
        var match = SignalFlowPlainLabelTightSpacingRegex.Match(candidate);
        if (!match.Success) {
            return segment;
        }

        return segment.Insert(start + match.Groups["label"].Length, " ");
    }

    private static bool LooksLikeListMarkerFragment(string value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        var leading = value.TrimStart();
        if (leading.Length == 0) {
            return false;
        }

        // Unordered list boundaries captured as right-side fragments after a newline.
        // Example captured values: "- ", "-  ", "-".
        if (leading[0] == '-' || leading[0] == '*' || leading[0] == '+') {
            if (leading.Length == 1) {
                return true;
            }

            var next = leading[1];
            if (char.IsWhiteSpace(next) || next == '*' || next == '`') {
                return true;
            }
        }

        var trimmed = leading.Trim();
        if (trimmed.Length < 2) {
            return false;
        }

        // Ordered list boundaries (for example, "2.** ...", "2) ...").
        int index = 0;
        while (index < trimmed.Length && char.IsDigit(trimmed[index])) {
            index++;
        }

        if (index == 0 || index != trimmed.Length - 1) {
            return false;
        }

        return trimmed[index] == '.' || trimmed[index] == ')';
    }
}
