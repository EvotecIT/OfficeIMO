namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private static bool FootnoteFragmentContinues(FootnoteFragment fragment) =>
            fragment.Start + fragment.Count < fragment.Note.Units.Count;

        // A new reference must retain its first legal fragment. Later references
        // keep their own minimum before an earlier note receives extra space.
        // This plan has no side effects, so notice reservation can safely refit it.
        private static bool TryPlanNewFootnoteFragments(IReadOnlyList<PreparedFootnote> notes,
            double available, out List<FootnoteFragment> fragments) {
            fragments = new List<FootnoteFragment>();
            int[] minimumCounts = notes.Select(note => GetMinimumFootnoteUnits(note.Units, 0)).ToArray();
            var minimumHeights = new double[notes.Count];
            double minimum = 0D;
            for (int index = 0; index < notes.Count; index++)
                minimum += minimumHeights[index] = MeasureFootnoteFragmentHeight(notes[index].Units, 0, minimumCounts[index]);
            if (notes.Count > 0 && minimum > available) {
                int last = notes.Count - 1;
                double clipped = MeasureFootnoteFragmentHeight(notes[last].Units, 0, minimumCounts[last],
                    Math.Max(0D, available - minimum + minimumHeights[last]));
                minimum -= minimumHeights[last] - clipped;
                minimumHeights[last] = clipped;
            }
            if (minimum > available + .001D) return false;
            for (int index = 0; index < notes.Count; index++) {
                PreparedFootnote note = notes[index];
                double laterMinimum = 0D;
                for (int next = index + 1; next < notes.Count; next++)
                    laterMinimum += minimumHeights[next];
                bool lastFragment = index == notes.Count - 1;
                int count = FitFootnoteUnits(note.Units, 0, available - laterMinimum, clipTrailingSpacing: lastFragment);
                if (count < minimumCounts[index])
                    throw new InvalidOperationException("Footnote admission lost its measured first fragment.");
                var fragment = new FootnoteFragment(note, 0, count, lastFragment ? available : double.PositiveInfinity);
                fragments.Add(fragment);
                available -= fragment.Height;
            }
            return true;
        }

        // Continuation pages consume notes in authored order. A later queued
        // note may wait for another page, but the first must make legal progress.
        private List<FootnoteFragment> PlanContinuedFootnoteFragments(double available) {
            var fragments = new List<FootnoteFragment>();
            foreach (FootnoteCursor cursor in continuedFootnotes) {
                int count = FitFootnoteUnits(cursor.Note.Units, cursor.Start, available, clipTrailingSpacing: true);
                if (count < GetMinimumFootnoteUnits(cursor.Note.Units, cursor.Start)) break;
                var fragment = new FootnoteFragment(cursor.Note, cursor.Start, count, available);
                fragments.Add(fragment);
                available -= fragment.Height;
                if (FootnoteFragmentContinues(fragment)) break;
            }
            return fragments;
        }
    }
}
