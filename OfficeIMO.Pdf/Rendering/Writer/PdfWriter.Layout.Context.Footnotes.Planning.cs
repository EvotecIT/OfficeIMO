namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private static bool FootnoteFragmentContinues(FootnoteFragment fragment) =>
            fragment.NextStart < fragment.Note.Units.Count;

        // Reserve each new reference's actual first legal fragment before greedily
        // spending room on an earlier note. A composite table keeps its accepted cut.
        private bool TryPlanNewFootnoteFragments(IReadOnlyList<PreparedFootnote> notes,
            double available, out List<FootnoteFragment> fragments) {
            fragments = new List<FootnoteFragment>();
            var minimumHeights = new double[notes.Count];
            double minimum = 0D;
            for (int index = 0; index < notes.Count; index++) {
                FootnoteFragment? first = FitFootnoteUnits(notes[index], 0, null, available,
                    clipTrailingSpacing: index == notes.Count - 1, minimumOnly: true);
                if (first == null) return false;
                minimum += minimumHeights[index] = first.Height;
            }
            if (notes.Count > 0 && minimum > available) {
                int last = notes.Count - 1;
                FootnoteFragment? clipped = FitFootnoteUnits(notes[last], 0, null,
                    Math.Max(0D, available - minimum + minimumHeights[last]), clipTrailingSpacing: true, minimumOnly: true);
                if (clipped != null) {
                    minimum -= minimumHeights[last] - clipped.Height;
                    minimumHeights[last] = clipped.Height;
                }
            }
            if (minimum > available + .001D) return false;
            for (int index = 0; index < notes.Count; index++) {
                double laterMinimum = 0D;
                for (int next = index + 1; next < notes.Count; next++) laterMinimum += minimumHeights[next];
                FootnoteFragment? fragment = FitFootnoteUnits(notes[index], 0, null, available - laterMinimum,
                    clipTrailingSpacing: index == notes.Count - 1);
                if (fragment == null) return false;
                fragments.Add(fragment);
                available -= fragment.Height;
            }
            return true;
        }

        private List<FootnoteFragment> PlanContinuedFootnoteFragments(double available) {
            var fragments = new List<FootnoteFragment>();
            foreach (FootnoteCursor cursor in continuedFootnotes) {
                FootnoteFragment? fragment = FitFootnoteUnits(cursor.Note, cursor.Start, cursor.Table, available,
                    clipTrailingSpacing: true);
                if (fragment == null) break;
                fragments.Add(fragment);
                available -= fragment.Height;
                if (FootnoteFragmentContinues(fragment)) break;
            }
            return fragments;
        }

        private FootnoteFragment? FitFootnoteUnits(PreparedFootnote note, int start, TableFlowCursor? tableCursor,
            double available, bool clipTrailingSpacing, bool minimumOnly = false) {
            NoteFlowFragment? flow = FitNoteFlowUnits(note.Units, start, tableCursor, available, clipTrailingSpacing, minimumOnly);
            return flow == null ? null : new FootnoteFragment(note, flow);
        }

        /// <summary>Accepted mixed note-flow content, shared by page notes and ordinary continued stories.</summary>
        private sealed class NoteFlowFragment {
            public NoteFlowFragment(int start, int count, double height, Dictionary<int, TableFragmentPlan> tables,
                TableFlowCursor? startTable = null) {
                Start = start; Count = count; Height = height;
                Tables = new System.Collections.ObjectModel.ReadOnlyDictionary<int, TableFragmentPlan>(tables);
                StartTable = startTable;
            }
            public int Start { get; }
            public int Count { get; }
            public double Height { get; }
            public System.Collections.ObjectModel.ReadOnlyDictionary<int, TableFragmentPlan> Tables { get; }
            public TableFlowCursor? StartTable { get; }
            public TableFlowCursor? NextTable => Tables.TryGetValue(Start + Count - 1, out var table) && table.Continues ? table.Next : null;
            public int NextStart => Start + Count - (NextTable == null ? 0 : 1);
        }

        /// <summary>Fits ordinary unit ranges and composite table fragments through one legal-cut dispatcher.</summary>
        private NoteFlowFragment? FitNoteFlowUnits(IReadOnlyList<FootnoteUnit> units, int start, TableFlowCursor? tableCursor,
            double available, bool clipTrailingSpacing, bool minimumOnly = false, bool relaxOversizedParagraphKeeps = false) {
            int minimum = GetMinimumFootnoteUnits(units, start, relaxOversizedParagraphKeeps: relaxOversizedParagraphKeeps);
            var tables = new Dictionary<int, TableFragmentPlan>();
            NoteFlowFragment? best = null;
            double maximum = clipTrailingSpacing ? available : double.PositiveInfinity;
            for (int end = start + 1; end <= units.Count; end++) {
                cancellationToken.ThrowIfCancellationRequested();
                FootnoteUnit unit = units[end - 1];
                if (unit.Table is { } geometry) {
                    TableFlowCursor? position = end - 1 == start ? tableCursor : null;
                    // A zero-height partial stand-in measures precisely the panel
                    // edges that must be withheld before asking the table owner.
                    var placeholderCursor = position ?? new TableFlowCursor(0, 0,
                        new TableSpanFlow(geometry.Table, geometry.Style, geometry.Columns,
                            geometry.Style.HeaderRowCount, geometry.Table.Rows.Count - geometry.Style.FooterRowCount));
                    tables[end - 1] = new TableFragmentPlan(geometry, placeholderCursor, placeholderCursor,
                        new List<TableRowSlice>(), 0D, 0D, 0D, new Dictionary<(int, int), int>(), 0D);
                    double overhead = MeasureFootnoteFragmentHeight(units, start, end - start,
                        double.PositiveInfinity, tables, tableCursor != null);
                    double full = Math.Max(0D, unit.MaximumKeepHeight -
                        MeasureFootnoteFragmentHeight(units, end - 1, 1, double.PositiveInfinity,
                            new Dictionary<int, TableFragmentPlan> { [end - 1] = tables[end - 1] }, position != null));
                    var completedCursor = new TableFlowCursor(geometry.Table.Rows.Count, 0, placeholderCursor.Spans);
                    tables[end - 1] = new TableFragmentPlan(geometry, placeholderCursor, completedCursor,
                        new List<TableRowSlice>(), 0D, 0D, 0D, new Dictionary<(int, int), int>(), 0D);
                    double finalOutset = Math.Max(0D, MeasureFootnoteFragmentHeight(units, start, end - start,
                        double.PositiveInfinity, tables, tableCursor != null) - overhead);
                    TableFragmentPlan? plan = PlanTableFragment(geometry, position, available - overhead, full,
                        minimumOnly && end - start >= minimum, finalOutset);
                    if (plan == null) break;
                    tables[end - 1] = plan;
                }
                int count = end - start;
                bool partial = tables.TryGetValue(end - 1, out var table) && table.Continues;
                double height = MeasureFootnoteFragmentHeight(units, start, count, maximum, tables, tableCursor != null);
                if (height > available + .001D) {
                    // A panel's boundary inset may be occupied by its next child.
                    // Only authored advances are monotone across candidate cuts.
                    double intrinsic = units[end - 1].PrefixHeight - (start == 0 ? 0D : units[start - 1].PrefixHeight);
                    foreach (var entry in tables) {
                        intrinsic += entry.Value.Height - units[entry.Key].Advance;
                        if (entry.Key == start && tableCursor != null) intrinsic -= units[entry.Key].PanelBefore;
                        if (entry.Key == end - 1 && partial) intrinsic -= units[entry.Key].PanelAfter;
                    }
                    double clippable = clipTrailingSpacing && !partial ? GetClippableFootnoteSpacing(units, end) : 0D;
                    if (partial || intrinsic - clippable > available + .001D) break;
                    continue;
                }
                if (count >= minimum && IsLegalFootnoteCut(units, start, count, relaxOversizedParagraphKeeps) &&
                    (!partial || CanCutFootnoteTable(units, start, end))) {
                    best = new NoteFlowFragment(start, count, height,
                        new Dictionary<int, TableFragmentPlan>(tables), tableCursor);
                    if (minimumOnly) return best;
                }
                if (partial) break;
            }
            return best;
        }

        private bool CanCutFootnoteTable(IReadOnlyList<FootnoteUnit> units, int start, int end) {
            foreach (FootnoteKeepRange range in units[end - 1].Keeps) {
                if (range.End > end && EnforceFootnoteKeep(units, range, start, true)) return false;
                if (!range.RelaxIfOversized && range.End >= end) return false;
            }
            return true;
        }
    }
}
