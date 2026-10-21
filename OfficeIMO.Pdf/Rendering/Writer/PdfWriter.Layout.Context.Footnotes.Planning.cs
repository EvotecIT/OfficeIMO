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
            double minimum = 0D;
            for (int index = 0; index < notes.Count; index++)
                minimum += notes[index].Units.Take(minimumCounts[index]).Sum(unit => unit.Height);
            if (minimum > available + .001D) return false;
            for (int index = 0; index < notes.Count; index++) {
                PreparedFootnote note = notes[index];
                double laterMinimum = 0D;
                for (int next = index + 1; next < notes.Count; next++)
                    laterMinimum += notes[next].Units.Take(minimumCounts[next]).Sum(unit => unit.Height);
                int count = FitFootnoteUnits(note.Units, 0, available - laterMinimum);
                if (count < minimumCounts[index])
                    throw new InvalidOperationException("Footnote admission lost its measured first fragment.");
                var fragment = new FootnoteFragment(note, 0, count);
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
                int count = FitFootnoteUnits(cursor.Note.Units, cursor.Start, available);
                if (count < GetMinimumFootnoteUnits(cursor.Note.Units, cursor.Start)) break;
                var fragment = new FootnoteFragment(cursor.Note, cursor.Start, count);
                fragments.Add(fragment);
                available -= fragment.Height;
                if (FootnoteFragmentContinues(fragment)) break;
            }
            return fragments;
        }
    }
}
