namespace OfficeIMO.Rtf;

/// <content>Provides rich semantic text and inline editing operations.</content>
public sealed partial class RtfParagraph {
    /// <summary>
    /// Replaces text across adjacent runs. Replacement text inherits formatting from the run where each match begins;
    /// unaffected prefix and suffix runs keep their existing formatting.
    /// </summary>
    public int ReplaceText(string oldText, string newText, StringComparison comparison = StringComparison.Ordinal) {
        if (oldText == null) throw new ArgumentNullException(nameof(oldText));
        if (newText == null) throw new ArgumentNullException(nameof(newText));
        if (oldText.Length == 0) throw new ArgumentException("Text to replace cannot be empty.", nameof(oldText));

        int replacements = 0;
        int inlineIndex = 0;
        while (inlineIndex < _inlines.Count) {
            if (!(_inlines[inlineIndex] is RtfRun)) {
                inlineIndex++;
                continue;
            }

            int segmentStart = inlineIndex;
            var runs = new List<RtfRun>();
            var starts = new List<int>();
            var text = new StringBuilder();
            while (inlineIndex < _inlines.Count && _inlines[inlineIndex] is RtfRun run) {
                starts.Add(text.Length);
                runs.Add(run);
                text.Append(run.Text);
                inlineIndex++;
            }

            var matches = new List<int>();
            int searchIndex = 0;
            string segmentText = text.ToString();
            while (searchIndex <= segmentText.Length - oldText.Length) {
                int match = segmentText.IndexOf(oldText, searchIndex, comparison);
                if (match < 0) break;
                matches.Add(match);
                searchIndex = match + oldText.Length;
            }

            for (int matchIndex = matches.Count - 1; matchIndex >= 0; matchIndex--) {
                int match = matches[matchIndex];
                int endOffset = match + oldText.Length;
                int startRunIndex = FindRunIndex(starts, runs, match);
                int endRunIndex = FindRunIndex(starts, runs, endOffset - 1);
                RtfRun startRun = runs[startRunIndex];
                RtfRun endRun = runs[endRunIndex];
                int startLocal = match - starts[startRunIndex];
                int endLocal = endOffset - starts[endRunIndex];

                if (startRunIndex == endRunIndex) {
                    startRun.Text = startRun.Text.Substring(0, startLocal) + newText + startRun.Text.Substring(endLocal);
                } else {
                    string suffix = endRun.Text.Substring(endLocal);
                    startRun.Text = startRun.Text.Substring(0, startLocal) + newText;
                    for (int removeIndex = endRunIndex - 1; removeIndex > startRunIndex; removeIndex--) {
                        RemoveInlineReference(runs[removeIndex]);
                    }

                    endRun.Text = suffix;
                    if (endRun.Text.Length == 0) RemoveInlineReference(endRun);
                }

                replacements++;
            }

            inlineIndex = segmentStart;
            while (inlineIndex < _inlines.Count && _inlines[inlineIndex] is RtfRun) inlineIndex++;
        }

        RebuildRuns();
        return replacements;
    }

    internal int FindBookmark(RtfBookmarkMarkerKind kind, string name) {
        for (int index = 0; index < _inlines.Count; index++) {
            if (_inlines[index] is RtfBookmarkMarker marker && marker.Kind == kind &&
                string.Equals(marker.Name, name, StringComparison.Ordinal)) return index;
        }

        return -1;
    }

    internal void ReplaceInlineRange(int startIndex, int count, string? replacement) {
        if (startIndex < 0 || startIndex > _inlines.Count) throw new ArgumentOutOfRangeException(nameof(startIndex));
        if (count < 0 || startIndex + count > _inlines.Count) throw new ArgumentOutOfRangeException(nameof(count));
        _inlines.RemoveRange(startIndex, count);
        if (!string.IsNullOrEmpty(replacement)) _inlines.Insert(startIndex, new RtfRun(replacement!));
        RebuildRuns();
    }

    private static int FindRunIndex(IReadOnlyList<int> starts, IReadOnlyList<RtfRun> runs, int offset) {
        for (int index = starts.Count - 1; index >= 0; index--) {
            if (offset >= starts[index] && offset < starts[index] + runs[index].Text.Length) return index;
        }

        return starts.Count - 1;
    }

    private void RemoveInlineReference(IRtfInline inline) {
        int index = _inlines.IndexOf(inline);
        if (index >= 0) _inlines.RemoveAt(index);
    }

    private void RebuildRuns() {
        _runs.Clear();
        foreach (IRtfInline inline in _inlines) {
            if (inline is RtfRun run) _runs.Add(run);
        }
    }
}
