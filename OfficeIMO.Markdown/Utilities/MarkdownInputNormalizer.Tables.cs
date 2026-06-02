using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownInputNormalizer {
    internal static string NormalizeCollapsedTableHeadingBoundaries(string? markdown) {
        return SplitCollapsedTableHeadingBoundaries(markdown ?? string.Empty);
    }

    private static string SplitCollapsedTableHeadingBoundaries(string value) {
        if (string.IsNullOrEmpty(value)) {
            return value;
        }

        var newline = value.IndexOf("\r\n", StringComparison.Ordinal) >= 0 ? "\r\n" : "\n";
        var lines = value.Replace("\r\n", "\n").Split('\n');
        if (lines.Length == 0) {
            return value;
        }

        var builder = new StringBuilder(value.Length + 16);
        bool inFence = false;
        char fenceChar = '\0';
        int fenceLength = 0;

        for (int i = 0; i < lines.Length; i++) {
            var line = lines[i] ?? string.Empty;
            var trimmedStart = line.TrimStart();

            if (TryReadFenceMarker(trimmedStart, out var currentFenceChar, out var currentFenceLength)) {
                if (!inFence) {
                    inFence = true;
                    fenceChar = currentFenceChar;
                    fenceLength = currentFenceLength;
                } else if (currentFenceChar == fenceChar && currentFenceLength >= fenceLength) {
                    inFence = false;
                    fenceChar = '\0';
                    fenceLength = 0;
                }

                builder.Append(line);
            } else if (!inFence && TrySplitCollapsedTableHeadingLine(line, out var left, out var right)) {
                builder.Append(left);
                builder.Append(newline);
                builder.Append(right);
            } else {
                builder.Append(line);
            }

            if (i < lines.Length - 1) {
                builder.Append(newline);
            }
        }

        return builder.ToString();
    }

    private static bool TryReadFenceMarker(string line, out char fenceChar, out int fenceLength) {
        fenceChar = '\0';
        fenceLength = 0;

        if (string.IsNullOrEmpty(line) || line.Length < 3) {
            return false;
        }

        char candidate = line[0];
        if (candidate != '`' && candidate != '~') {
            return false;
        }

        int run = 1;
        while (run < line.Length && line[run] == candidate) {
            run++;
        }

        if (run < 3) {
            return false;
        }

        fenceChar = candidate;
        fenceLength = run;
        return true;
    }

    private static bool TrySplitCollapsedTableHeadingLine(string line, out string left, out string right) {
        left = string.Empty;
        right = string.Empty;

        if (string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        var trimmedStart = line.TrimStart();
        if (!trimmedStart.StartsWith("|", StringComparison.Ordinal)) {
            return false;
        }

        int pipeCount = 0;
        int headingIndex = -1;
        int codeFenceLength = 0;

        for (int i = 0; i < line.Length; i++) {
            char ch = line[i];

            if (ch == '\\' && i + 1 < line.Length) {
                i++;
                continue;
            }

            if (ch == '`') {
                int run = 1;
                while (i + run < line.Length && line[i + run] == '`') {
                    run++;
                }

                if (codeFenceLength == 0) {
                    codeFenceLength = run;
                } else if (run == codeFenceLength) {
                    codeFenceLength = 0;
                }

                i += run - 1;
                continue;
            }

            if (codeFenceLength != 0) {
                continue;
            }

            if (ch == '|') {
                pipeCount++;
                continue;
            }

            if (ch != '#' || pipeCount < 2) {
                continue;
            }

            int runLength = 1;
            while (i + runLength < line.Length && line[i + runLength] == '#') {
                runLength++;
            }

            if (runLength > 6) {
                continue;
            }

            int afterMarker = i + runLength;
            if (afterMarker >= line.Length || !char.IsWhiteSpace(line[afterMarker])) {
                continue;
            }

            headingIndex = i;
            break;
        }

        if (headingIndex <= 0) {
            return false;
        }

        left = line.Substring(0, headingIndex).TrimEnd();
        right = line.Substring(headingIndex).TrimStart();
        return left.Length > 0 && right.Length > 0;
    }
}
