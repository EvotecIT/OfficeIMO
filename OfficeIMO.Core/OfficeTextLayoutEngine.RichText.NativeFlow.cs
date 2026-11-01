using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextLayoutEngine {
    private static OfficeTextTabPlacement? FindNativeTab(IReadOnlyList<OfficeTextTabPlacement>? tabs, int sourceEnd) {
        int left = 0, right = (tabs?.Count ?? 0) - 1;
        while (left <= right) {
            int middle = left + (right - left) / 2;
            var tab = tabs![middle];
            if (tab.SourceEnd == sourceEnd) return tab;
            if (tab.SourceEnd > sourceEnd) right = middle - 1;
            else left = middle + 1;
        }
        return null;
    }

    private static bool ContainsSourceEnd(IReadOnlyList<int> lineEnds, int sourceEnd) {
        int left = 0, right = lineEnds.Count - 1;
        while (left <= right) {
            int middle = left + (right - left) / 2;
            int end = lineEnds[middle];
            if (end == sourceEnd) return true;
            if (end > sourceEnd) right = middle - 1;
            else left = middle + 1;
        }
        return false;
    }

    /// <summary>Repaints a selected native region without rewrapping its shortened final word or inserting source characters.</summary>
    internal static OfficeRichTextBlockLayout LayoutNativeFlowLines(IReadOnlyList<OfficeRichTextRun> runs,
        double width, double lineHeightFactor, OfficeRichTextMeasurement measure,
        OfficeTextParagraphIndent indent, OfficeTextTabStops? tabs, OfficeTextTabLeaderLayout.Budget leaderBudget,
        OfficeRichTextFlowPlan plan, double scale, CancellationToken token) {
        var normalized = NormalizeRichTextRuns(runs, out bool truncated, token);
        int length = 0;
        foreach (var run in normalized) length += run.Text.Length;
        var lineEnds = new List<int>(); var hyphens = new List<int>();
        foreach (var line in plan.Lines) {
            if (line.SourceEnd < length) lineEnds.Add(line.SourceEnd);
            if (line.GeneratedHyphen) hyphens.Add(line.SourceEnd);
        }
        var layout = LayoutRichTextBlockCore(normalized, width, double.MaxValue, lineHeightFactor, measure,
            false, OfficeTextOverflowBehavior.Clip, indent, truncated, token,
            hardBreakStartsParagraph: false, tabStops: tabs, leaderBudget: leaderBudget, nativeFlowLineEnds: lineEnds,
            nativeFlowTabs: plan.Tabs, nativeFlowScale: scale, nativeGeneratedHyphens: hyphens);
        for (int i = 0; i < layout.Lines.Count && i < plan.Lines.Count; i++) {
            if (plan.Lines[i].FittingWidth is double fittingWidth) {
                var line = layout.Lines[i];
                line.FittingBounds = (line.OffsetX, line.OffsetX + fittingWidth * scale);
            }
        }
        return layout;
    }

    private static IEnumerable<RichTextToken> ApplyNativeFlowLineEnds(IEnumerable<RichTextToken> tokens,
        IReadOnlyList<int> lineEnds, IReadOnlyList<int>? hyphens, CancellationToken cancellationToken) {
        int index = 0, lastEnd = 0; OfficeRichTextRun? lastRun = null;
        foreach (var token in tokens) {
            cancellationToken.ThrowIfCancellationRequested();
            lastEnd = token.SourceEnd; lastRun = token.Run;
            if (token.HardBreak) {
                while (index < lineEnds.Count && lineEnds[index] <= token.SourceEnd) index++;
                yield return token;
                continue;
            }
            int start = token.SourceStart;
            while (index < lineEnds.Count && lineEnds[index] <= token.SourceEnd) {
                int end = lineEnds[index++];
                if (end <= start) continue;
                yield return RichTextToken.CreateText(token.Run,
                    token.Text.Substring(start - token.SourceStart, end - start), token.IsWhitespace, start, end);
                if (hyphens != null && ContainsSourceEnd(hyphens, end)) yield return RichTextToken.CreateGeneratedHyphen(token.Run, end);
                yield return RichTextToken.CreateHardBreak(token.Run, end);
                start = end;
            }
            if (start < token.SourceEnd) yield return start == token.SourceStart ? token : RichTextToken.CreateText(token.Run,
                token.Text.Substring(start - token.SourceStart), token.IsWhitespace, start, token.SourceEnd);
        }
        if (lastRun != null && hyphens != null && ContainsSourceEnd(hyphens, lastEnd)
            && !ContainsSourceEnd(lineEnds, lastEnd)) yield return RichTextToken.CreateGeneratedHyphen(lastRun, lastEnd);
    }
}
