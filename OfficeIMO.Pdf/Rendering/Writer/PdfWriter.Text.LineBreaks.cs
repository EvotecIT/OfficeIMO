using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private readonly struct PdfTextTokenChunk {
        public PdfTextTokenChunk(string text, double width) {
            Text = text;
            Width = width;
        }

        public string Text { get; }

        public double Width { get; }
    }

    private static System.Collections.Generic.List<PdfTextTokenChunk>? TryBuildMultilingualTokenChunks(
        string token,
        Func<string, double> measure,
        double firstMaxWidth,
        double subsequentMaxWidth) {
        return TryBuildTokenChunks(token, GetMultilingualBreakpoints(token), measure, firstMaxWidth, subsequentMaxWidth);
    }

    private static System.Collections.Generic.List<PdfTextTokenChunk>? TryBuildSoftLineBreakTokenChunks(
        string token,
        PdfOptions? options,
        Func<string, double> measure,
        double firstMaxWidth,
        double subsequentMaxWidth) {
        return TryBuildTokenChunks(token, GetValidSoftLineBreakpoints(token, options), measure, firstMaxWidth, subsequentMaxWidth, allowFinalOverflow: true);
    }

    private static System.Collections.Generic.List<PdfTextTokenChunk>? TryBuildTokenChunks(
        string token,
        int[] breakpoints,
        Func<string, double> measure,
        double firstMaxWidth,
        double subsequentMaxWidth,
        bool allowFinalOverflow = false) {
        if (breakpoints.Length == 0) {
            return null;
        }

        var chunks = new System.Collections.Generic.List<PdfTextTokenChunk>();
        int position = 0;
        while (position < token.Length) {
            int selectedBreak = -1;
            string selectedText = string.Empty;
            double selectedWidth = 0D;
            double currentMaxWidth = Math.Max(1D, chunks.Count == 0 ? firstMaxWidth : subsequentMaxWidth);

            foreach (int candidate in EnumerateTokenBreakCandidates(breakpoints, position, token.Length)) {
                string chunkText = token.Substring(position, candidate - position);
                if (chunkText.Length == 0) {
                    continue;
                }

                double chunkWidth = measure(chunkText);
                if (chunkWidth <= currentMaxWidth) {
                    selectedBreak = candidate;
                    selectedText = chunkText;
                    selectedWidth = chunkWidth;
                } else if (allowFinalOverflow && candidate == token.Length && chunks.Count > 0 && selectedBreak < 0) {
                    selectedBreak = candidate;
                    selectedText = chunkText;
                    selectedWidth = chunkWidth;
                } else if (selectedBreak >= 0) {
                    break;
                }
            }

            if (selectedBreak <= position || selectedText.Length == 0) {
                return null;
            }

            chunks.Add(new PdfTextTokenChunk(selectedText, selectedWidth));
            position = selectedBreak;
        }

        return chunks.Count > 1 ? chunks : null;
    }

    private static int[] GetValidSoftLineBreakpoints(string token, PdfOptions? options) {
        PdfTextLineBreakCallback? callback = options?.TextLineBreakCallbackSnapshot;
        if (callback == null || string.IsNullOrEmpty(token)) {
            return Array.Empty<int>();
        }

        System.Collections.Generic.IReadOnlyList<int>? points = callback(token);
        if (points == null || points.Count == 0) {
            return Array.Empty<int>();
        }

        return points
            .Where(point => IsValidTokenBreakIndex(token, point))
            .Distinct()
            .OrderBy(point => point)
            .ToArray();
    }

    private static System.Collections.Generic.IEnumerable<int> EnumerateTokenBreakCandidates(int[] breakpoints, int position, int tokenLength) {
        for (int index = 0; index < breakpoints.Length; index++) {
            int point = breakpoints[index];
            if (point > position) {
                yield return point;
            }
        }

        yield return tokenLength;
    }

    private static int[] GetMultilingualBreakpoints(string token) {
        if (string.IsNullOrEmpty(token)) {
            return Array.Empty<int>();
        }

        var breakpoints = new System.Collections.Generic.List<int>();
        int leftIndex = 0;
        int left = ReadScalar(token, ref leftIndex);
        while (leftIndex < token.Length) {
            int boundary = leftIndex;
            int rightIndex = leftIndex;
            int right = ReadScalar(token, ref rightIndex);
            if (IsValidTokenBreakIndex(token, boundary) && CanBreakBetweenScalars(left, right)) {
                breakpoints.Add(boundary);
            }

            left = right;
            leftIndex = rightIndex;
        }

        return breakpoints.ToArray();
    }

    private static bool CanBreakBetweenScalars(int left, int right) {
        if (!IsCjkLineBreakScalar(left) && !IsCjkLineBreakScalar(right)) {
            return false;
        }

        if (IsLineBreakNonStarter(right) || IsLineBreakOpeningPunctuation(left) || IsLineBreakClosingPunctuation(right)) {
            return false;
        }

        return true;
    }

    private static bool IsLineBreakNonStarter(int scalar) {
        if (scalar == 0x200D ||
            (scalar >= 0xFE00 && scalar <= 0xFE0F) ||
            (scalar >= 0xE0100 && scalar <= 0xE01EF)) {
            return true;
        }

        UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(char.ConvertFromUtf32(scalar), 0);
        return category == UnicodeCategory.NonSpacingMark ||
            category == UnicodeCategory.SpacingCombiningMark ||
            category == UnicodeCategory.EnclosingMark;
    }

    private static bool IsCjkLineBreakScalar(int scalar) =>
        (scalar >= 0x3040 && scalar <= 0x309F) ||
        (scalar >= 0x30A0 && scalar <= 0x30FF) ||
        (scalar >= 0x31F0 && scalar <= 0x31FF) ||
        (scalar >= 0x3400 && scalar <= 0x4DBF) ||
        (scalar >= 0x4E00 && scalar <= 0x9FFF) ||
        (scalar >= 0xAC00 && scalar <= 0xD7AF) ||
        (scalar >= 0xF900 && scalar <= 0xFAFF) ||
        (scalar >= 0x20000 && scalar <= 0x2FA1F);

    private static bool IsLineBreakOpeningPunctuation(int scalar) => scalar switch {
        '(' or '[' or '{' or 0x201C or 0x2018 or 0x3008 or 0x300A or 0x300C or 0x300E or 0x3010 or 0x3014 or 0x3016 or 0x3018 or 0x301A or 0xFF08 or 0xFF3B or 0xFF5B => true,
        _ => false
    };

    private static bool IsLineBreakClosingPunctuation(int scalar) => scalar switch {
        ')' or ']' or '}' or ',' or '.' or ':' or ';' or '!' or '?' or 0x2019 or 0x201D or 0x2026 or 0x3001 or 0x3002 or 0x3009 or 0x300B or 0x300D or 0x300F or 0x3011 or 0x3015 or 0x3017 or 0x3019 or 0x301B or 0xFF01 or 0xFF09 or 0xFF0C or 0xFF0E or 0xFF1A or 0xFF1B or 0xFF1F or 0xFF3D or 0xFF5D => true,
        _ => false
    };
}
