namespace OfficeIMO.Diff;

/// <summary>
/// Slides pure insertions and deletions to the position where they start and end at separators.
/// </summary>
/// <remarks>
/// When the item before a change equals its last item (or the item after it equals its first), the same change can be
/// placed in several positions with the same cost; "a. b c." minus "b c." can drop ". b c" or "b c.". The engine picks
/// one arbitrarily. This pass picks the position whose edges touch separators (whitespace for text), so whole words and
/// sentences are reported. The edit distance never changes.
/// </remarks>
internal static class OfficeDiffBoundaries {
    private const int MaxSlide = 256;

    internal static IReadOnlyList<OfficeDiffMatch> Align(IReadOnlyList<OfficeDiffMatch> source, int[] oldIds, int[] newIds, bool[] oldSeparators, bool[] newSeparators) {
        var matches = new List<OfficeDiffMatch>(source);
        for (int index = 1; index < matches.Count; index++) {
            OfficeDiffMatch before = matches[index - 1];
            OfficeDiffMatch after = matches[index];
            int oldStart = before.OldStart + before.Length, oldEnd = after.OldStart;
            int newStart = before.NewStart + before.Length, newEnd = after.NewStart;
            bool deletion = oldEnd > oldStart && newEnd == newStart;
            bool insertion = newEnd > newStart && oldEnd == oldStart;
            if (!deletion && !insertion) continue;

            int[] ids = deletion ? oldIds : newIds;
            bool[] separators = deletion ? oldSeparators : newSeparators;
            int start = deletion ? oldStart : newStart;
            int end = deletion ? oldEnd : newEnd;

            int left = 0;
            while (left < before.Length && left < MaxSlide && ids[start - left - 1] == ids[end - left - 1]) left++;
            int right = 0;
            while (right < after.Length && right < MaxSlide && ids[start + right] == ids[end + right]) right++;
            if (left == 0 && right == 0) continue;

            int bestShift = 0;
            int bestScore = Score(separators, start, end);
            for (int shift = -left; shift <= right; shift++) {
                if (shift == 0) continue;
                int score = Score(separators, start + shift, end + shift);
                if (score > bestScore || (score == bestScore && Math.Abs(shift) < Math.Abs(bestShift))) {
                    bestScore = score;
                    bestShift = shift;
                }
            }

            if (bestShift == 0) continue;
            matches[index - 1] = new OfficeDiffMatch(before.OldStart, before.NewStart, before.Length + bestShift);
            matches[index] = new OfficeDiffMatch(after.OldStart + bestShift, after.NewStart + bestShift, after.Length - bestShift);
        }

        matches.RemoveAll(static match => match.Length == 0);
        return matches;
    }

    // Two points for each edge of the change that touches a separator, inside or just outside it; one more for each
    // separator just outside, and one when the change starts with content, so "Drop that. " beats " Drop that.".
    private static int Score(bool[] separators, int start, int end) {
        int score = 0;
        if (start == 0 || separators[start - 1] || separators[start]) score += 2;
        if (end == separators.Length || separators[end] || separators[end - 1]) score += 2;
        if (start == 0 || separators[start - 1]) score++;
        if (end == separators.Length || separators[end]) score++;
        if (!separators[start]) score++;
        return score;
    }
}
