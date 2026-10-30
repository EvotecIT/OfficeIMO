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
        PdfCreateOptions? options,
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
        bool allowFinalOverflow = false,
        Func<string, double>? measureFirstChunk = null,
        bool appendHyphen = false,
        bool allowSingleChunk = false,
        System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (breakpoints.Length == 0) {
            return null;
        }

        var chunks = new System.Collections.Generic.List<PdfTextTokenChunk>();
        long remainingMeasurementWork = Math.Max(4096L, 64L * token.Length);
        int position = 0;
        while (position < token.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int selectedBreak = -1;
            string selectedText = string.Empty;
            double selectedWidth = 0D;
            double currentMaxWidth = Math.Max(0D, chunks.Count == 0 ? firstMaxWidth : subsequentMaxWidth);

            foreach (int candidate in EnumerateTokenBreakCandidates(breakpoints, position, token.Length)) {
                cancellationToken.ThrowIfCancellationRequested();
                long length = candidate - position + (appendHyphen && candidate < token.Length ? 1 : 0);
                long cost = length * (chunks.Count == 0 && measureFirstChunk != null ? 2 : 1);
                if (cost > remainingMeasurementWork) return null;
                remainingMeasurementWork -= cost;
                string chunkText = token.Substring(position, candidate - position);
                if (appendHyphen && candidate < token.Length) chunkText += "-";
                if (chunkText.Length == 0) {
                    continue;
                }

                double chunkWidth = measure(chunkText);
                double fitWidth = chunks.Count == 0 && measureFirstChunk != null ? measureFirstChunk(chunkText) : chunkWidth;
                if (fitWidth <= currentMaxWidth) {
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

        return chunks.Count > 1 || allowSingleChunk ? chunks : null;
    }

    private static int[] GetValidSoftLineBreakpoints(string token, PdfCreateOptions? options) {
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


    // Largest prefix ending at a UAX #14 opportunity (or a hyphenation point, which adds a hyphen)
    // that fits the space left on the line.
    private static bool TrySplitTokenAtLineEnd(string token, double available, PdfCreateOptions? options, Func<string, double> measure,
        out string head, out double headWidth, out int consumed,
        System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        head = string.Empty;
        headWidth = 0D;
        consumed = 0;
        if (available <= 0D) return false;
        var candidates = new SortedDictionary<int, bool>();
        foreach (int position in OfficeTextLineBreaks.GetBreakPositions(token, allowCjkBreaks: true, cancellationToken)) candidates[position] = false;
        foreach (int position in GetValidHyphenationBreakpoints(token, options)) {
            if (!candidates.ContainsKey(position)) candidates[position] = true;
        }
        bool identifier = IsHyphenatedIdentifier(token);
        // Try short prefixes first. On long CJK or zero-width tokens, measuring every
        // overlapping prefix would make one line's layout quadratic in token length.
        long remainingMeasurementWork = Math.Max(4096L, 32L * token.Length);
        foreach (KeyValuePair<int, bool> candidate in candidates) {
            cancellationToken.ThrowIfCancellationRequested();
            if (candidate.Key <= 0 || candidate.Key >= token.Length) continue;
            if (identifier && !candidate.Value && IsHyphenCharacter(token[candidate.Key - 1])) continue;
            int length = candidate.Key + (candidate.Value ? 1 : 0);
            if (length > remainingMeasurementWork) break;
            remainingMeasurementWork -= length;
            string text = token.Substring(0, candidate.Key);
            if (candidate.Value) text += "-";
            double width = measure(text);
            if (width > available) continue;
            head = text;
            headWidth = width;
            consumed = candidate.Key;
        }
        return consumed > 0;
    }

    private static bool IsHyphenatedIdentifier(string token) {
        bool hasDigit = false, hasHyphen = false;
        foreach (char ch in token) {
            if (char.IsDigit(ch)) hasDigit = true;
            else if (IsHyphenCharacter(ch)) hasHyphen = true;
            else if (!char.IsLetter(ch) && ch != '.' && ch != '/' && ch != '_' && ch != '#') return false;
        }
        return hasDigit && hasHyphen;
    }

    private static bool IsHyphenCharacter(char ch) => ch == '-' || ch == '\u2010' || ch == '\u2011';
}
