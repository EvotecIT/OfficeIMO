using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownInputNormalizer {
    private static string FlattenNestedStrongSpansOutsideFencedCodeBlocks(string value) {
        return MarkdownFence.ApplyTransformOutsideFencedCodeBlocks(
            value,
            static segment => MarkdownInlineCode.ApplyTransformPreservingInlineCodeSpans(segment, FlattenNestedStrongSpansPreservingInlineCode));
    }

    private static string FlattenNestedStrongSpansPreservingInlineCode(string input) {
        var current = input ?? string.Empty;
        while (true) {
            var flattened = NestedStrongSpanRegex.Replace(
                current,
                static match =>
                    "**"
                    + match.Groups["left"].Value
                    + match.Groups["inner"].Value
                    + match.Groups["right"].Value
                    + "**");
            if (flattened.Equals(current, StringComparison.Ordinal)) {
                break;
            }

            current = flattened;
        }

        return FlattenLabeledOuterStrongSpans(current);
    }

    private static string FlattenLabeledOuterStrongSpans(string input) {
        if (string.IsNullOrEmpty(input) || input.IndexOf("**", StringComparison.Ordinal) < 0) {
            return input ?? string.Empty;
        }

        return LabeledOuterStrongLineRegex.Replace(input, static match => {
            var body = match.Groups["body"].Value;
            if (body.IndexOf("**", StringComparison.Ordinal) < 0) {
                return match.Value;
            }

            var trimmedBody = body.TrimEnd();
            if (trimmedBody.Length == 0) {
                return match.Value;
            }

            var lastBodyChar = trimmedBody[trimmedBody.Length - 1];
            if (lastBodyChar != '.' && lastBodyChar != '!' && lastBodyChar != '?' && lastBodyChar != ')') {
                return match.Value;
            }

            var cleaned = FlattenNestedStrongMarkers(body);
            if (cleaned.Equals(body, StringComparison.Ordinal)) {
                return match.Value;
            }

            return match.Groups["prefix"].Value + cleaned + match.Groups["suffix"].Value + match.Groups["tail"].Value;
        });
    }

    private static string FlattenNestedStrongMarkers(string input) {
        if (string.IsNullOrEmpty(input) || input.IndexOf("**", StringComparison.Ordinal) < 0) {
            return input ?? string.Empty;
        }

        var current = input;
        for (var i = 0; i < StrongFlattenMaxIterations; i++) {
            var next = SimpleNestedStrongSpanRegex.Replace(
                current,
                match => {
                    var inner = match.Groups["inner"].Value;
                    if (inner.Length == 0) {
                        return inner;
                    }

                    var prefix = string.Empty;
                    var suffix = string.Empty;
                    var start = match.Index;
                    var end = match.Index + match.Length;
                    if (start > 0) {
                        var before = current[start - 1];
                        if (!char.IsWhiteSpace(before) && IsWordLikeChar(before) && IsWordLikeChar(inner[0])) {
                            prefix = " ";
                        }
                    }

                    if (end < current.Length) {
                        var after = current[end];
                        if (!char.IsWhiteSpace(after) && IsWordLikeChar(inner[inner.Length - 1]) && IsWordLikeChar(after)) {
                            suffix = " ";
                        }
                    }

                    return prefix + inner + suffix;
                });

            if (next.Equals(current, StringComparison.Ordinal)) {
                return next;
            }

            current = next;
        }

        return current;
    }

    private static bool IsWordLikeChar(char value) {
        return char.IsLetterOrDigit(value);
    }

    private static string RepairBrokenTwoLineStrongLeadIns(string text) {
        if (string.IsNullOrEmpty(text) || text.IndexOf("**", StringComparison.Ordinal) < 0) {
            return text ?? string.Empty;
        }

        var hasCrLf = text.IndexOf("\r\n", StringComparison.Ordinal) >= 0;
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        var rewritten = new List<string>(lines.Length);
        var changed = false;

        for (var i = 0; i < lines.Length; i++) {
            var current = lines[i] ?? string.Empty;
            if (i + 1 < lines.Length) {
                var currentMatch = BrokenTwoLineStrongLeadInRegex.Match(current);
                if (currentMatch.Success) {
                    var next = lines[i + 1] ?? string.Empty;
                    var closingIndex = next.IndexOf("**", StringComparison.Ordinal);
                    if (closingIndex > 0) {
                        var label = currentMatch.Groups["label"].Value.Trim().TrimEnd(':');
                        var body = next.Substring(0, closingIndex).Trim();
                        var tail = next.Substring(closingIndex + 2).Trim();
                        if (label.Length > 0
                            && body.Length > 0
                            && !StructuralMarkdownLineRegex.IsMatch(body)) {
                            var merged = currentMatch.Groups["indent"].Value
                                         + "**" + label + ":** "
                                         + body
                                         + (tail.Length == 0 ? string.Empty : " " + tail);
                            rewritten.Add(merged);
                            changed = true;
                            i++;
                            continue;
                        }
                    }
                }
            }

            rewritten.Add(current);
        }

        if (!changed) {
            return text;
        }

        var rebuilt = string.Join("\n", rewritten);
        return hasCrLf ? rebuilt.Replace("\n", "\r\n") : rebuilt;
    }

    private static string RepairDanglingTrailingStrongListClosers(string text) {
        if (string.IsNullOrEmpty(text) || text.IndexOf("****", StringComparison.Ordinal) < 0) {
            return text ?? string.Empty;
        }

        var hasCrLf = text.IndexOf("\r\n", StringComparison.Ordinal) >= 0;
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        var changed = false;

        for (var i = 0; i < lines.Length; i++) {
            var line = lines[i] ?? string.Empty;
            var trimmedStart = line.TrimStart();
            if (!trimmedStart.StartsWith("- ", StringComparison.Ordinal)
                && !OrderedListLeadRegex.IsMatch(trimmedStart)) {
                continue;
            }

            var repaired = TrailingDanglingStrongListTokenRegex.Replace(line, static match => {
                var token = match.Groups["token"].Value.Trim();
                if (token.Length == 0 || token.IndexOf("**", StringComparison.Ordinal) >= 0) {
                    return match.Value;
                }

                return "**" + token + "**" + match.Groups["tail"].Value;
            });

            if (repaired.Equals(line, StringComparison.Ordinal)) {
                continue;
            }

            lines[i] = repaired;
            changed = true;
        }

        if (!changed) {
            return text;
        }

        var rebuilt = string.Join("\n", lines);
        return hasCrLf ? rebuilt.Replace("\n", "\r\n") : rebuilt;
    }
}
