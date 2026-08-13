namespace OfficeIMO.Pdf;

internal static partial class PdfTextEditor {
    private static int[]? TryBuildAuthoredBoundaryMap(string normalized, string authored) {
        var boundaries = new int[normalized.Length + 1];
        int normalizedIndex = 0;
        int authoredIndex = 0;
        while (normalizedIndex < normalized.Length && authoredIndex < authored.Length) {
            boundaries[normalizedIndex] = authoredIndex;
            if (char.IsWhiteSpace(normalized[normalizedIndex]) && char.IsWhiteSpace(authored[authoredIndex])) {
                int normalizedEnd = normalizedIndex + 1;
                while (normalizedEnd < normalized.Length && char.IsWhiteSpace(normalized[normalizedEnd])) normalizedEnd++;
                int authoredEnd = authoredIndex + 1;
                while (authoredEnd < authored.Length && char.IsWhiteSpace(authored[authoredEnd])) authoredEnd++;
                for (int index = normalizedIndex + 1; index < normalizedEnd; index++) {
                    boundaries[index] = authoredIndex + (int)((long)(authoredEnd - authoredIndex) * (index - normalizedIndex) / (normalizedEnd - normalizedIndex));
                }
                normalizedIndex = normalizedEnd;
                authoredIndex = authoredEnd;
                boundaries[normalizedIndex] = authoredIndex;
                continue;
            }
            if (normalized[normalizedIndex] != authored[authoredIndex]) return null;
            normalizedIndex++;
            authoredIndex++;
            boundaries[normalizedIndex] = authoredIndex;
        }
        if (normalizedIndex != normalized.Length || authoredIndex != authored.Length) return null;
        return boundaries;
    }

    private static int[] BuildIdentityBoundaryMap(int length) {
        var boundaries = new int[length + 1];
        for (int index = 0; index <= length; index++) boundaries[index] = index;
        return boundaries;
    }

    private static string TrimAuthoredEdgeWhitespace(string value) {
        int start = 0;
        while (start < value.Length && char.IsWhiteSpace(value[start])) start++;
        int end = value.Length;
        while (end > start && char.IsWhiteSpace(value[end - 1])) end--;
        return value.Substring(start, end - start);
    }
}
