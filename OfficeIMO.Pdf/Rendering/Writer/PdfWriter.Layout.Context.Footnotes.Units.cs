namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private sealed class FootnoteUnit {
            public IPdfBlock? Block { get; init; }
            public PreparedTableGeometry? Table { get; init; }
            public List<RichSeg>? Line { get; init; }
            public double Advance { get; init; }
            public double Before { get; init; }
            public double After { get; init; }
            public double Size { get; init; }
            public double Leading { get; init; }
            public double X { get; init; }
            public double Width { get; init; }
            public double FrameX { get; init; }
            public double FrameWidth { get; init; }
            public int ParagraphLine { get; init; }
            public int ParagraphLines { get; init; }
            public int MinimumOrphanLines { get; init; } = 1;
            public int MinimumWidowLines { get; init; } = 1;
            public bool KeepParagraphTogether { get; init; }
            public FootnotePanel[] Panels { get; init; } = Array.Empty<FootnotePanel>();
            public List<FootnoteKeepRange> Keeps { get; } = new();
            public PreparedListLayout? List { get; init; }
            public PreparedListItem? ListItem { get; init; }
            public double PanelBefore { get; set; }
            public double PanelAfter { get; set; }
            public double PrefixHeight { get; set; }
            public double MaximumKeepHeight { get; set; } = double.PositiveInfinity;
            public double Height => PanelBefore + Before + Advance + After + PanelAfter;
            public bool IsVisual => Line != null || Table != null || Block is ImageBlock or HorizontalRuleBlock;
        }

        private sealed class FootnoteKeepRange {
            public FootnoteKeepRange(int start, int end, bool relaxIfOversized) {
                Start = start; End = end; RelaxIfOversized = relaxIfOversized;
            }
            public int Start { get; }
            public int End { get; }
            public bool RelaxIfOversized { get; }
        }

        private sealed class FootnoteBlockSpan {
            public FootnoteBlockSpan(int start, int end, bool keepWithNext, bool marker) {
                Start = start; End = end; KeepWithNext = keepWithNext; IsMarker = marker;
            }
            public int Start { get; }
            public int End { get; }
            public bool KeepWithNext { get; }
            public bool IsMarker { get; }
        }

        private static void AddFootnoteKeepRange(List<FootnoteUnit> units, int start, int end, bool relaxIfOversized) {
            if (end <= start) return;
            var range = new FootnoteKeepRange(start, end, relaxIfOversized);
            for (int index = start; index < end; index++) units[index].Keeps.Add(range);
        }

        // A chain retains each kept block and only the terminal block's first
        // legal fragment. Consolidating adjacent links avoids quadratic ranges.
        private void AddFootnoteKeepChains(List<FootnoteUnit> units, List<FootnoteBlockSpan> spans) {
            for (int index = 0; index < spans.Count; index++) {
                if (!spans[index].KeepWithNext) continue;
                int next = index + 1;
                int inspected = 1;
                while (next < spans.Count && spans[next].IsMarker) next++;
                if (next == spans.Count) continue;
                while (spans[next].KeepWithNext) {
                    if (++inspected > MaxKeepWithNextChainBlocks)
                        throw new NotSupportedException("KeepWithNext chains cannot contain more than " + MaxKeepWithNextChainBlocks + " visual blocks.");
                    int following = next + 1;
                    while (following < spans.Count && spans[following].IsMarker) following++;
                    if (following == spans.Count) break;
                    next = following;
                }
                int end = spans[next].Start + GetMinimumFootnoteUnits(units, spans[next].Start, checkCapacity: false);
                AddFootnoteKeepRange(units, spans[index].Start, end, relaxIfOversized: true);
                index = next - 1;
            }
        }

        private static void CompleteFootnoteUnitMeasurements(List<FootnoteUnit> units) {
            double height = 0D;
            foreach (FootnoteUnit unit in units) {
                height += unit.Height;
                unit.PrefixHeight = height;
            }
        }

        private static double MeasureFootnoteUnits(IReadOnlyList<FootnoteUnit> units) =>
            MeasureFootnoteFragmentHeight(units, 0, units.Count);

        private static double MeasureFootnoteFragmentHeight(IReadOnlyList<FootnoteUnit> units, int start, int count,
            double maximumHeight = double.PositiveInfinity, IReadOnlyDictionary<int, TableFragmentPlan>? tables = null,
            bool continuedTable = false) {
            if (count == 0) return 0D;
            int end = start + count;
            double height = units[end - 1].PrefixHeight - (start == 0 ? 0D : units[start - 1].PrefixHeight);
            bool partialTable = tables != null && tables.TryGetValue(end - 1, out var lastTable) && lastTable.Continues;
            if (tables != null) {
                foreach (var entry in tables) {
                    FootnoteUnit unit = units[entry.Key];
                    height += entry.Value.Height - unit.Advance;
                    if (entry.Key == start && continuedTable) height -= unit.PanelBefore;
                    if (entry.Key == end - 1 && partialTable) height -= unit.PanelAfter;
                }
            }
            foreach (FootnotePanel panel in units[start].Panels)
                if (panel.Start < start || continuedTable && panel.Start == start) height += panel.GetTopPadding(true, tables);
            foreach (FootnotePanel panel in units[end - 1].Panels) {
                bool continues = panel.End > end || partialTable && panel.End == end;
                double bottom = panel.Style.GetFragmentBottomPadding(continues);
                if (continues) height += bottom;
                height += Math.Max(0D, panel.Style.FragmentBottomInset - bottom);
            }
            return height - Math.Min(partialTable ? 0D : GetClippableFootnoteSpacing(units, end), Math.Max(0D, height - maximumHeight));
        }

        private static double GetClippableFootnoteSpacing(IReadOnlyList<FootnoteUnit> units, int end) =>
            units[end - 1].Panels.Where(panel => panel.End == end && panel.Style.ClipSpacingAfterAtFrameEnd)
                .Sum(panel => panel.Style.SpacingAfter);

        private bool EnforceFootnoteKeep(IReadOnlyList<FootnoteUnit> units, FootnoteKeepRange range, int start, bool checkCapacity) {
            if (!range.RelaxIfOversized || !checkCapacity) return true;
            int remainingStart = Math.Max(start, range.Start);
            double height = MeasureFootnoteFragmentHeight(units, remainingStart, range.End - remainingStart);
            FootnoteUnit terminal = units[range.End - 1];
            if (terminal.Table is { } table) {
                // A keep chain's terminal table contributes its first legal
                // fragment. Earlier kept tables still contribute their whole box.
                double frame = Math.Max(0D, terminal.MaximumKeepHeight - terminal.PanelBefore - terminal.PanelAfter);
                TableFragmentPlan? first = PlanTableFragment(table, null, frame, frame, minimumOnly: true);
                if (first != null) height += first.Height - terminal.Advance;
            }
            return height <= units[remainingStart].MaximumKeepHeight + .001D;
        }

        // A contained table cell follows the ordinary cell owner's paragraph
        // policy: retain a fitting paragraph, but allow an oversized paragraph
        // to make progress. Other note/story callers keep their existing policy.
        private static bool EnforceContainedParagraphKeep(IReadOnlyList<FootnoteUnit> units, int index,
            bool relaxOversizedParagraphKeeps) {
            FootnoteUnit unit = units[index];
            if (!unit.KeepParagraphTogether) return false;
            if (!relaxOversizedParagraphKeeps) return true;
            int first = index - unit.ParagraphLine;
            double height = MeasureFootnoteFragmentHeight(units, first, unit.ParagraphLines);
            return height <= unit.MaximumKeepHeight + .001D;
        }

        private int GetMinimumFootnoteUnits(IReadOnlyList<FootnoteUnit> units, int start, bool checkCapacity = true,
            bool relaxOversizedParagraphKeeps = false) {
            int end = start;
            while (end < units.Count) {
                FootnoteUnit unit = units[end++];
                if (unit.Line != null) {
                    int remaining = unit.ParagraphLines - unit.ParagraphLine;
                    int required = EnforceContainedParagraphKeep(units, end - 1, relaxOversizedParagraphKeeps && checkCapacity) ? remaining : unit.MinimumOrphanLines;
                    if (remaining < required + unit.MinimumWidowLines) required = remaining;
                    end += Math.Max(0, Math.Min(required, remaining) - 1);
                    break;
                }
                if (unit.IsVisual) break;
            }
            for (int index = start; index < end; index++)
                foreach (FootnoteKeepRange range in units[index].Keeps)
                    if (range.End > end && EnforceFootnoteKeep(units, range, start, checkCapacity)) end = range.End;
            return end - start;
        }

        private bool IsLegalFootnoteCut(IReadOnlyList<FootnoteUnit> units, int start, int count, bool relaxOversizedParagraphKeeps = false) {
            int end = start + count;
            FootnoteUnit last = units[end - 1];
            if (last.Line != null && last.ParagraphLine < last.ParagraphLines - 1) {
                if (EnforceContainedParagraphKeep(units, end - 1, relaxOversizedParagraphKeeps) ||
                    last.ParagraphLines - last.ParagraphLine - 1 < last.MinimumWidowLines) return false;
                int first = Math.Max(start, end - last.ParagraphLine - 1);
                if (end - first < units[first].MinimumOrphanLines) return false;
            }
            foreach (FootnoteKeepRange range in last.Keeps)
                if (range.End > end && EnforceFootnoteKeep(units, range, start, checkCapacity: true)) return false;
            return true;
        }

    }
}
