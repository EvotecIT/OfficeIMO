using System.Globalization;
using System.Text;

namespace OfficeIMO.Diff;

/// <summary>Writes line diffs in the unified format read by <c>patch</c> and <c>git apply</c>.</summary>
public static class OfficeUnifiedDiff {
    /// <summary>Compares two texts line by line and writes a unified diff. Returns an empty string when the texts are equal.</summary>
    /// <param name="oldText">Old text.</param>
    /// <param name="newText">New text.</param>
    /// <param name="oldName">Name on the <c>---</c> header line.</param>
    /// <param name="newName">Name on the <c>+++</c> header line.</param>
    /// <param name="contextLines">Unchanged lines shown around each change.</param>
    /// <param name="options">Diff settings.</param>
    /// <param name="cancellationToken">Cancels the alignment.</param>
    public static string Create(string? oldText, string? newText, string oldName = "a", string newName = "b", int contextLines = 3,
        OfficeDiffOptions? options = null, CancellationToken cancellationToken = default) =>
        Format(OfficeDiff.CompareLines(oldText, newText, options, cancellationToken), oldName, newName, contextLines);

    /// <summary>Writes a unified diff for a line diff produced by <see cref="OfficeDiff.CompareLines"/>.</summary>
    public static string Format(OfficeDiffResult<OfficeTextToken> lineDiff, string oldName = "a", string newName = "b", int contextLines = 3) {
        if (lineDiff is null) throw new ArgumentNullException(nameof(lineDiff));
        if (contextLines < 0) throw new ArgumentOutOfRangeException(nameof(contextLines));
        if (lineDiff.AreEqual) return string.Empty;

        var output = new StringBuilder();
        output.Append("--- ").Append(oldName ?? "a").Append('\n');
        output.Append("+++ ").Append(newName ?? "b").Append('\n');

        IReadOnlyList<OfficeDiffHunk> regions = lineDiff.GetChangedRegions();
        int regionIndex = 0;
        while (regionIndex < regions.Count) {
            int last = regionIndex;
            while (last + 1 < regions.Count &&
                   regions[last + 1].OldStart - (regions[last].OldStart + regions[last].OldLength) <= contextLines * 2) {
                last++;
            }

            int oldStart = Math.Max(0, regions[regionIndex].OldStart - contextLines);
            int newStart = Math.Max(0, regions[regionIndex].NewStart - contextLines);
            int oldEnd = Math.Min(lineDiff.Old.Count, regions[last].OldStart + regions[last].OldLength + contextLines);
            int newEnd = Math.Min(lineDiff.New.Count, regions[last].NewStart + regions[last].NewLength + contextLines);
            output.Append("@@ -").Append(Range(oldStart, oldEnd - oldStart)).Append(" +").Append(Range(newStart, newEnd - newStart)).Append(" @@\n");

            int oldCursor = oldStart;
            int newCursor = newStart;
            for (int index = regionIndex; index <= last; index++) {
                OfficeDiffHunk region = regions[index];
                while (oldCursor < region.OldStart) {
                    AppendLine(output, ' ', lineDiff.Old[oldCursor].Text);
                    oldCursor++;
                    newCursor++;
                }

                for (int offset = 0; offset < region.OldLength; offset++) AppendLine(output, '-', lineDiff.Old[region.OldStart + offset].Text);
                for (int offset = 0; offset < region.NewLength; offset++) AppendLine(output, '+', lineDiff.New[region.NewStart + offset].Text);
                oldCursor = region.OldStart + region.OldLength;
                newCursor = region.NewStart + region.NewLength;
            }

            while (oldCursor < oldEnd) {
                AppendLine(output, ' ', lineDiff.Old[oldCursor].Text);
                oldCursor++;
            }

            regionIndex = last + 1;
        }

        return output.ToString();
    }

    private static string Range(int start, int length) {
        // GNU diff writes the line before an empty range, and omits a count of one.
        if (length == 0) return start.ToString(CultureInfo.InvariantCulture) + ",0";
        string first = (start + 1).ToString(CultureInfo.InvariantCulture);
        return length == 1 ? first : first + "," + length.ToString(CultureInfo.InvariantCulture);
    }

    private static void AppendLine(StringBuilder output, char marker, string line) {
        output.Append(marker);
        bool terminated = line.Length > 0 && (line[line.Length - 1] == '\n' || line[line.Length - 1] == '\r');
        if (line.Length > 0 && line[line.Length - 1] == '\n') output.Append(line);
        else if (terminated) output.Append(line, 0, line.Length - 1).Append('\n');
        else output.Append(line).Append('\n');
        if (!terminated) output.Append("\\ No newline at end of file\n");
    }
}
