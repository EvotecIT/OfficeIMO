using OfficeIMO.Drawing;

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
        return TryBuildTokenChunks(token, OfficeTextLineBreaks.GetBreakPositions(token).ToArray(), measure, firstMaxWidth, subsequentMaxWidth);
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
        Func<string, IReadOnlyList<int>>? callback = options?.TextLineBreakCallbackSnapshot;
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

}
