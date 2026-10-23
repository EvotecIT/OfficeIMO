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
        private IReadOnlyList<FootnoteUnit> pageFootnoteContinuationNotice = Array.Empty<FootnoteUnit>();
        private bool pageReserveContinuationNoticeSpaceOnFinalPage;
        private bool pageFootnotesContinue;

        // Page coordinates: the body must end at or above this Y. Running-footer exclusion
        // is already part of currentOpts; the footnote plan never mutates those options.
        private double FootnoteBodyBottom => currentOpts.MarginBottom + (paintingFootnotes ? 0D : pageFootnoteHeight);
        private bool HasPageFootnotes => pageFootnotes.Count > 0;

        private sealed class FootnoteCursor {
            public FootnoteCursor(PreparedFootnote note, int start, TableFlowCursor? table = null) { Note = note; Start = start; Table = table; }
            public PreparedFootnote Note { get; }
            public int Start { get; }
            public TableFlowCursor? Table { get; }
        }

        private sealed class FootnoteFragment {
            public FootnoteFragment(PreparedFootnote note, NoteFlowFragment flow) { Note = note; Flow = flow; }
            public PreparedFootnote Note { get; }
            public NoteFlowFragment Flow { get; }
            public int Start => Flow.Start;
            public int Count => Flow.Count;
            public double Height => Flow.Height;
            public System.Collections.ObjectModel.ReadOnlyDictionary<int, TableFragmentPlan> Tables => Flow.Tables;
            public TableFlowCursor? StartTable => Flow.StartTable;
            public TableFlowCursor? NextTable => Flow.NextTable;
            public int NextStart => Flow.NextStart;
        }

        private sealed class PreparedFootnote {
            public PreparedFootnote(PdfFootnoteDefinition definition, List<FootnoteUnit> units,
                List<FootnoteUnit> separator, List<FootnoteUnit> continuationSeparator,
                List<FootnoteUnit> continuationNotice, double usableWidth) {
                Definition = definition; Units = units; Separator = separator; ContinuationSeparator = continuationSeparator;
                ContinuationNotice = continuationNotice;
                UsableWidth = usableWidth;
            }
            public PdfFootnoteDefinition Definition { get; }
            public List<FootnoteUnit> Units { get; }
            public List<FootnoteUnit> Separator { get; }
            public List<FootnoteUnit> ContinuationSeparator { get; }
            public List<FootnoteUnit> ContinuationNotice { get; }
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
            IReadOnlyList<FootnoteUnit> separator = pageFootnotes.Count == 0 ? notes[0].Separator : Array.Empty<FootnoteUnit>();
            double separatorHeight = MeasureFootnoteUnits(separator);
            IReadOnlyList<FootnoteUnit> notice = HasPageFootnotes ? pageFootnoteContinuationNotice : notes[0].ContinuationNotice;
            bool reserveFinalNotice = HasPageFootnotes ? pageReserveContinuationNoticeSpaceOnFinalPage
                : notes[0].Definition.ReserveContinuationNoticeSpaceOnFinalPage;
            double noticeHeight = MeasureFootnoteUnits(notice);
            double oldNoticeHeight = pageFootnotesContinue || pageReserveContinuationNoticeSpaceOnFinalPage
                ? MeasureFootnoteUnits(pageFootnoteContinuationNotice) : 0D;
            double available = bodyBottom - currentOpts.MarginBottom - (pageFootnoteHeight - oldNoticeHeight) - separatorHeight;
            bool noticeAlreadyReserved = pageFootnotesContinue || reserveFinalNotice;
            if (!TryPlanNewFootnoteFragments(notes, available - (noticeAlreadyReserved ? noticeHeight : 0D), out var fragments)) return false;
            bool continues = pageFootnotesContinue || fragments.Any(FootnoteFragmentContinues);
            // Ordinary notes reserve notice space only after finding continuation.
            // A captured final-page policy has already withheld the measured flow.
            if (continues && !noticeAlreadyReserved && noticeHeight > 0D &&
                !TryPlanNewFootnoteFragments(notes, available - noticeHeight, out fragments)) return false;
            reservedHeight = pageFootnoteHeight - oldNoticeHeight + separatorHeight +
                fragments.Sum(fragment => fragment.Height) + (continues || reserveFinalNotice ? noticeHeight : 0D);
            if (commit) {
                pageFootnotes.AddRange(fragments);
                foreach (PreparedFootnote note in notes) scheduledFootnotes.Add(note.Definition);
                if (pageFootnoteSeparator.Count == 0 && separator.Count > 0) pageFootnoteSeparator = separator;
                pageFootnoteContinuationNotice = notice;
                pageReserveContinuationNoticeSpaceOnFinalPage = reserveFinalNotice;
                pageFootnotesContinue = continues;
                pageFootnoteHeight = reservedHeight;
            }
            return true;
        }

        private void StartPageFootnotes() {
            footnotePageOptions = currentOpts;
            pageFootnotes.Clear();
            pageFootnoteSeparator = Array.Empty<FootnoteUnit>();
            pageFootnoteContinuationNotice = Array.Empty<FootnoteUnit>();
            pageReserveContinuationNoticeSpaceOnFinalPage = false;
            pageFootnotesContinue = false;
            pageFootnoteHeight = 0D;
            if (continuedFootnotes.Count == 0) return;
            foreach (FootnoteCursor cursor in continuedFootnotes) {
                PdfDocument.ValidateFootnoteBlocks(cursor.Note.Definition.Blocks, currentOpts);
                PdfDocument.ValidateFootnoteBlocks(cursor.Note.Definition.Separator, currentOpts, allowTables: false);
                PdfDocument.ValidateFootnoteBlocks(cursor.Note.Definition.ContinuationSeparator, currentOpts, allowTables: false);
                PdfDocument.ValidateFootnoteBlocks(cursor.Note.Definition.ContinuationNotice, currentOpts, allowTables: false);
            }
            double usableWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
            if (continuedFootnotes.Any(cursor => Math.Abs(cursor.Note.UsableWidth - usableWidth) > .001D))
                throw new NotSupportedException("A continued footnote cannot cross a change in usable page width. Finish the note before changing page geometry.");
            pageFootnoteSeparator = continuedFootnotes[0].Note.ContinuationSeparator;
            pageFootnoteContinuationNotice = continuedFootnotes[0].Note.ContinuationNotice;
            pageReserveContinuationNoticeSpaceOnFinalPage = continuedFootnotes[0].Note.Definition.ReserveContinuationNoticeSpaceOnFinalPage;
            double separatorHeight = MeasureFootnoteUnits(pageFootnoteSeparator);
            double noticeHeight = MeasureFootnoteUnits(pageFootnoteContinuationNotice);
            double available = yStart - currentOpts.MarginBottom - separatorHeight;
            List<FootnoteFragment> fragments = PlanContinuedFootnoteFragments(available -
                (pageReserveContinuationNoticeSpaceOnFinalPage ? noticeHeight : 0D));
            pageFootnotesContinue = fragments.Count < continuedFootnotes.Count || fragments.Any(FootnoteFragmentContinues);
            if (pageFootnotesContinue && !pageReserveContinuationNoticeSpaceOnFinalPage && noticeHeight > 0D)
                fragments = PlanContinuedFootnoteFragments(available - noticeHeight);
            if (fragments.Count == 0)
                throw new ArgumentException("The first continued footnote fragment and its separator or notice exceed the available page body height.");
            pageFootnotes.AddRange(fragments);
            pageFootnoteHeight = separatorHeight + fragments.Sum(fragment => fragment.Height) +
                (pageFootnotesContinue || pageReserveContinuationNoticeSpaceOnFinalPage ? noticeHeight : 0D);
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
                PaintFootnoteUnits(pageFootnoteSeparator, 0, pageFootnoteSeparator.Count, ref top);
                foreach (FootnoteFragment fragment in pageFootnotes)
                    PaintFootnoteUnits(fragment.Note.Units, fragment.Start, fragment.Count, ref top, fragment.Height, fragment.Tables, fragment.StartTable != null);
                if (pageFootnotesContinue)
                    PaintFootnoteUnits(pageFootnoteContinuationNotice, 0, pageFootnoteContinuationNotice.Count, ref top);
            } finally {
                paintingFootnotes = false;
                currentOpts = savedOptions;
                y = savedY;
                pageHasBodyContent = savedBodyContent;
            }
            var remaining = new List<FootnoteCursor>();
            foreach (FootnoteFragment fragment in pageFootnotes) {
                int next = fragment.NextStart;
                if (next < fragment.Note.Units.Count) remaining.Add(new FootnoteCursor(fragment.Note, next, fragment.NextTable));
            }
            foreach (FootnoteCursor cursor in continuedFootnotes)
                if (!pageFootnotes.Any(fragment => ReferenceEquals(fragment.Note, cursor.Note))) remaining.Add(cursor);
            continuedFootnotes.Clear();
            continuedFootnotes.AddRange(remaining);
            pageFootnotes.Clear();
            pageFootnoteSeparator = Array.Empty<FootnoteUnit>();
            pageFootnoteContinuationNotice = Array.Empty<FootnoteUnit>();
            pageReserveContinuationNoticeSpaceOnFinalPage = false;
            pageFootnotesContinue = false;
            pageFootnoteHeight = 0D;
        }

        // An authored page boundary follows all fragments associated with the
        // preceding body. Automatic pagination still allows body to resume on
        // the last continued-note page. Use NewPage so active containers retain
        // their normal fragment preparation and resumption.
        private void NewPageAfterFootnoteContinuations(bool preserveEmptyPage = false) {
            NewPage(preserveEmptyPage);
            while (HasPageFootnotes) NewPage();
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
