namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private bool HasUnscheduledFootnote(IEnumerable<RichSeg> segments) =>
            segments.Any(segment => segment.FootnoteReference != null && !scheduledFootnotes.Contains(segment.FootnoteReference));

        // Paragraph after-space can overflow the ordinary physical frame, but it
        // must not enter an accepted or newly introduced page-note region. Keep
        // enclosing-container padding and its already included after-space once.
        private double GetFootnoteClosingTextPadding(IReadOnlyList<List<RichSeg>> lines, int start,
            double spacingAfter, bool closesContainer = true) {
            double closing = closesContainer ? GetClosingTextPadding(spacingAfter) : 0D;
            if (paintingFootnotes || spacingAfter <= 0D || closing >= spacingAfter) return closing;
            return HasPageFootnotes || HasUnscheduledFootnote(lines.Skip(start).SelectMany(line => line))
                ? spacingAfter : closing;
        }

        /// <summary>Ends a body slice at its first newly admitted reference, so its note has priority over later body text.</summary>
        private bool ConstrainFootnoteSlice(IReadOnlyList<List<RichSeg>> lines, List<double> heights,
            int start, int minimumLines, int minimumTailLines, double before, double closingPadding,
            ref int take, ref double heightSum) {
            int end = start + take;
            for (int index = start; index < end; index++) {
                if (!HasUnscheduledFootnote(lines[index])) continue;
                int mandatoryEnd = Math.Min(lines.Count, Math.Max(index + 1, start + Math.Max(1, minimumLines)));
                if (lines.Count - mandatoryEnd < minimumTailLines) mandatoryEnd = lines.Count;
                mandatoryEnd = Math.Min(mandatoryEnd, end);
                double advance = 0D;
                for (int line = start; line < mandatoryEnd; line++) advance += heights[line];
                double bottom = y - before - advance - (mandatoryEnd == lines.Count ? closingPadding : 0D);
                if (!TryReserveFootnotes(lines, start, mandatoryEnd - start, bottom, commit: false)) {
                    // Paint only body preceding the reference. The reference and its first legal
                    // note fragment move together; previews have not consumed either identity.
                    take = minimumLines >= lines.Count - start ? 0 : index - start;
                    heightSum = 0D;
                    for (int line = start; line < start + take; line++) heightSum += heights[line];
                    return take > 0;
                }
                take = mandatoryEnd - start;
                heightSum = advance;
                return true;
            }
            return true;
        }

        private void RequireFootnoteSlice(IReadOnlyList<List<RichSeg>> lines, int start, int count, double bodyBottom) {
            if (!TryReserveFootnotes(lines, start, count, bodyBottom, commit: true))
                throw new ArgumentException("The reference and first footnote fragment exceed the available page body height.");
        }
    }
}
