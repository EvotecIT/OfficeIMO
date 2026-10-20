namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly HashSet<PdfFootnoteDefinition> scheduledFootnotes = new();
        private readonly Dictionary<PdfFootnoteDefinition, int> footnoteSourceOrdinals = new();
        private readonly List<FootnoteCursor> continuedFootnotes = new();
        private readonly List<FootnoteFragment> pageFootnotes = new();
        private readonly Dictionary<(PdfFootnoteDefinition Definition, PdfOptions Options, double Width), PreparedFootnote> preparedFootnotes = new();
        private bool paintingFootnotes;
        private double pageFootnoteHeight;
        private PdfOptions? footnotePageOptions;
        private IReadOnlyList<FootnoteUnit> pageFootnoteSeparator = Array.Empty<FootnoteUnit>();

        // Page coordinates: the body must end at or above this Y. Running-footer exclusion
        // is already part of currentOpts; the footnote plan never mutates those options.
        private double FootnoteBodyBottom => currentOpts.MarginBottom + (paintingFootnotes ? 0D : pageFootnoteHeight);
        private bool HasPageFootnotes => pageFootnotes.Count > 0;

        private sealed class FootnoteCursor {
            public FootnoteCursor(PreparedFootnote note, int start) { Note = note; Start = start; }
            public PreparedFootnote Note { get; }
            public int Start { get; }
        }

        private sealed class FootnoteFragment {
            public FootnoteFragment(PreparedFootnote note, int start, int count) {
                Note = note; Start = start; Count = count;
                Height = note.Units.Skip(start).Take(count).Sum(unit => unit.Height);
            }
            public PreparedFootnote Note { get; }
            public int Start { get; }
            public int Count { get; }
            public double Height { get; }
        }

        private sealed class PreparedFootnote {
            public PreparedFootnote(PdfFootnoteDefinition definition, List<FootnoteUnit> units,
                List<FootnoteUnit> separator, List<FootnoteUnit> continuationSeparator, double usableWidth) {
                Definition = definition; Units = units; Separator = separator; ContinuationSeparator = continuationSeparator;
                UsableWidth = usableWidth;
            }
            public PdfFootnoteDefinition Definition { get; }
            public List<FootnoteUnit> Units { get; }
            public List<FootnoteUnit> Separator { get; }
            public List<FootnoteUnit> ContinuationSeparator { get; }
            public double UsableWidth { get; }
        }

        /// <summary>Assigns identity order from authored body references, independently of definition creation and fragment admission.</summary>
        private void RegisterFootnoteSourceOrder(IReadOnlyList<IPdfBlock> blocks) {
            foreach (IPdfBlock block in blocks) {
                cancellationToken.ThrowIfCancellationRequested();
                foreach (PdfFootnoteDefinition definition in PdfFootnoteTraversal.References(block))
                    if (!footnoteSourceOrdinals.ContainsKey(definition)) footnoteSourceOrdinals.Add(definition, footnoteSourceOrdinals.Count);
            }
        }

        private int FootnoteSourceOrdinal(PdfFootnoteDefinition definition) =>
            footnoteSourceOrdinals.TryGetValue(definition, out int ordinal) ? ordinal : int.MaxValue;

        private bool TryReserveFootnotes(IReadOnlyList<List<RichSeg>> lines, int start, int count, double bodyBottom, bool commit) =>
            TryReserveFootnotes(lines.Skip(start).Take(count).SelectMany(line => line), bodyBottom, commit);

        /// <summary>Preflights or accepts references carried by the actual candidate fragment. bodyBottom is a page Y coordinate.</summary>
        private bool TryReserveFootnotes(IEnumerable<RichSeg> segments, double bodyBottom, bool commit) =>
            TryReserveFootnotes(segments, bodyBottom, commit, out _);

        private bool TryReserveFootnotes(IEnumerable<RichSeg> segments, double bodyBottom, bool commit, out double reservedHeight) {
            cancellationToken.ThrowIfCancellationRequested();
            reservedHeight = pageFootnoteHeight;
            var definitions = new List<PdfFootnoteDefinition>();
            foreach (RichSeg segment in segments) {
                if (segment.FootnoteReference is not { } definition || scheduledFootnotes.Contains(definition) || definitions.Contains(definition)) continue;
                if (isRunningContent || paintingFootnotes || activeColumnFlow != null || HasFloatingTables)
                    throw new NotSupportedException("Page-associated footnotes require ordinary single-column body flow.");
                definitions.Add(definition);
            }
            // This helper adds footnote admission only. With no new association or
            // accepted page plan, ordinary flow retains its own overlap/overflow rules.
            if (definitions.Count == 0 && !HasPageFootnotes) return true;
            if (bodyBottom < FootnoteBodyBottom - .001D) return false;
            if (definitions.Count == 0) return true;
            definitions.Sort((left, right) => FootnoteSourceOrdinal(left).CompareTo(FootnoteSourceOrdinal(right)));
            var notes = definitions.Select(PrepareFootnote).ToList();
            double available = bodyBottom - currentOpts.MarginBottom - pageFootnoteHeight;
            IReadOnlyList<FootnoteUnit> separator = pageFootnotes.Count == 0 ? notes[0].Separator : Array.Empty<FootnoteUnit>();
            double separatorHeight = separator.Sum(unit => unit.Height);
            int[] minimumCounts = notes.Select(note => GetMinimumFootnoteUnits(note.Units, 0)).ToArray();
            double minimum = separatorHeight;
            for (int index = 0; index < notes.Count; index++) minimum += notes[index].Units.Take(minimumCounts[index]).Sum(unit => unit.Height);
            if (minimum > available + .001D) return false;

            // The reference's mandatory body fragment wins first. The largest legal note prefix
            // then wins over following body; later references retain their first legal fragments.
            available -= separatorHeight;
            reservedHeight += separatorHeight;
            for (int index = 0; index < notes.Count; index++) {
                PreparedFootnote note = notes[index];
                double laterMinimum = 0D;
                for (int next = index + 1; next < notes.Count; next++)
                    laterMinimum += notes[next].Units.Take(minimumCounts[next]).Sum(unit => unit.Height);
                int count = FitFootnoteUnits(note.Units, 0, available - laterMinimum);
                if (count < minimumCounts[index]) throw new InvalidOperationException("Footnote admission lost its measured first fragment.");
                var fragment = new FootnoteFragment(note, 0, count);
                reservedHeight += fragment.Height;
                if (commit) pageFootnotes.Add(fragment);
                available -= note.Units.Take(count).Sum(unit => unit.Height);
                if (commit) scheduledFootnotes.Add(note.Definition);
            }
            if (commit) {
                if (pageFootnoteSeparator.Count == 0 && separator.Count > 0) pageFootnoteSeparator = separator;
                pageFootnoteHeight = reservedHeight;
            }
            return true;
        }

        private static int GetMinimumFootnoteUnits(IReadOnlyList<FootnoteUnit> units, int start) {
            int count = 0;
            while (start + count < units.Count) {
                FootnoteUnit unit = units[start + count++];
                if (unit.Line != null) {
                    int required = unit.KeepParagraphTogether ? unit.ParagraphLines - unit.ParagraphLine : unit.MinimumOrphanLines;
                    if (unit.ParagraphLines - unit.ParagraphLine < required + unit.MinimumWidowLines)
                        required = unit.ParagraphLines - unit.ParagraphLine;
                    count += Math.Max(0, Math.Min(required, unit.ParagraphLines - unit.ParagraphLine) - 1);
                    break;
                }
                if (unit.IsVisual) break;
            }
            return count;
        }

        private static int FitFootnoteUnits(IReadOnlyList<FootnoteUnit> units, int start, double available) {
            int count = 0;
            for (int index = start; index < units.Count; index++) {
                if (units[index].Height > available + .001D) break;
                available -= units[index].Height;
                count++;
            }
            while (count > 0) {
                FootnoteUnit last = units[start + count - 1];
                if (last.Line == null || last.ParagraphLine == last.ParagraphLines - 1) break;
                int tail = last.ParagraphLines - last.ParagraphLine - 1;
                if (!last.KeepParagraphTogether && tail >= last.MinimumWidowLines) break;
                count--;
            }
            if (count < GetMinimumFootnoteUnits(units, start)) return 0;
            return count;
        }

        private void StartPageFootnotes() {
            footnotePageOptions = currentOpts;
            pageFootnotes.Clear();
            pageFootnoteSeparator = Array.Empty<FootnoteUnit>();
            pageFootnoteHeight = 0D;
            if (continuedFootnotes.Count == 0) return;
            double usableWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
            if (continuedFootnotes.Any(cursor => Math.Abs(cursor.Note.UsableWidth - usableWidth) > .001D))
                throw new NotSupportedException("A continued footnote cannot cross a change in usable page width. Finish the note before changing page geometry.");
            double available = yStart - currentOpts.MarginBottom;
            for (int index = 0; index < continuedFootnotes.Count; index++) {
                FootnoteCursor cursor = continuedFootnotes[index];
                IReadOnlyList<FootnoteUnit> separator = index == 0 ? cursor.Note.ContinuationSeparator : Array.Empty<FootnoteUnit>();
                double separatorHeight = separator.Sum(unit => unit.Height);
                int count = FitFootnoteUnits(cursor.Note.Units, cursor.Start, available - separatorHeight);
                if (count < GetMinimumFootnoteUnits(cursor.Note.Units, cursor.Start)) {
                    if (pageFootnotes.Count == 0)
                        throw new ArgumentException("The first continued footnote fragment exceeds the available page body height.");
                    break;
                }
                var fragment = new FootnoteFragment(cursor.Note, cursor.Start, count);
                pageFootnotes.Add(fragment);
                if (index == 0) pageFootnoteSeparator = separator;
                pageFootnoteHeight += fragment.Height + separatorHeight;
                available -= fragment.Height + separatorHeight;
                if (cursor.Start + count < cursor.Note.Units.Count) break;
            }
        }

        private void FlushPageFootnotes() {
            if (pageFootnotes.Count == 0) return;
            double savedY = y;
            bool savedBodyContent = pageHasBodyContent;
            bool beneath = pageFootnotes.All(fragment => fragment.Note.Definition.Placement == PdfFootnotePlacement.BeneathText);
            PdfOptions savedOptions = currentOpts;
            currentOpts = footnotePageOptions ?? currentOpts;
            double top = beneath ? y : currentOpts.MarginBottom + pageFootnoteHeight;
            pageFootnotes.Sort((left, right) => FootnoteSourceOrdinal(left.Note.Definition).CompareTo(FootnoteSourceOrdinal(right.Note.Definition)));
            paintingFootnotes = true;
            try {
                foreach (FootnoteUnit unit in pageFootnoteSeparator) PaintFootnoteUnit(unit, ref top);
                foreach (FootnoteFragment fragment in pageFootnotes) {
                    for (int index = fragment.Start; index < fragment.Start + fragment.Count; index++)
                        PaintFootnoteUnit(fragment.Note.Units[index], ref top);
                }
            } finally {
                paintingFootnotes = false;
                currentOpts = savedOptions;
                y = savedY;
                pageHasBodyContent = savedBodyContent;
            }
            var remaining = new List<FootnoteCursor>();
            foreach (FootnoteFragment fragment in pageFootnotes) {
                int next = fragment.Start + fragment.Count;
                if (next < fragment.Note.Units.Count) remaining.Add(new FootnoteCursor(fragment.Note, next));
            }
            foreach (FootnoteCursor cursor in continuedFootnotes)
                if (!pageFootnotes.Any(fragment => ReferenceEquals(fragment.Note, cursor.Note))) remaining.Add(cursor);
            continuedFootnotes.Clear();
            continuedFootnotes.AddRange(remaining);
            pageFootnotes.Clear();
            pageFootnoteSeparator = Array.Empty<FootnoteUnit>();
            pageFootnoteHeight = 0D;
        }

        private void FinishFootnoteContinuations() {
            while (continuedFootnotes.Count > 0 || HasPageFootnotes) {
                cancellationToken.ThrowIfCancellationRequested();
                if (currentPage != null) FlushPage(force: true);
                if (continuedFootnotes.Count == 0) break;
                StartPage(currentPageBaseOptions);
            }
        }
    }
}
