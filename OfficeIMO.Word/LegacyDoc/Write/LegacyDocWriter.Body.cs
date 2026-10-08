using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private readonly struct LegacyDocWritableBody {
            internal LegacyDocWritableBody(
                string text,
                IReadOnlyList<LegacyDocWritableRun> formattedRuns,
                IReadOnlyList<LegacyDocWritableParagraph> formattedParagraphs,
                LegacyDocWritableBookmarks bookmarks,
                IReadOnlyList<LegacyDocWritableSection> sections,
                LegacyDocWritableStyleSheet styleSheet,
                Numbering? numbering,
                LegacyDocWritableFootnoteStories footnoteStories,
                LegacyDocWritableEndnoteStories endnoteStories,
                LegacyDocWritableHeaderFooterStories headerFooterStories,
                LegacyDocWritableCommentStories commentStories,
                byte[] pictureData,
                bool hasPictures,
                bool facingPages,
                bool mirrorMargins,
                bool gutterAtTop,
                bool noColumnBalance,
                ushort defaultTabStop,
                EndnotePositionValues? endnotePosition,
                bool trackRevisions,
                bool lockRevisionTracking) {
                ChpxPages = Array.Empty<IReadOnlyList<LegacyDocWritableSegment>>();
                PapxPages = Array.Empty<IReadOnlyList<LegacyDocWritableParagraphSegment>>();
                Text = text;
                FormattedRuns = formattedRuns;
                FormattedParagraphs = formattedParagraphs;
                Sections = sections;
                StyleSheet = styleSheet;
                FootnoteText = footnoteStories.Text;
                PlcffndRef = footnoteStories.PlcffndRef;
                PlcffndTxt = footnoteStories.PlcffndTxt;
                FootnoteMarkerPositions = footnoteStories.MarkerPositions;
                FootnoteFormattedRuns = footnoteStories.FormattedRuns;
                FootnoteFormattedParagraphs = footnoteStories.FormattedParagraphs;
                EndnoteText = endnoteStories.Text;
                PlcfendRef = endnoteStories.PlcfendRef;
                PlcfendTxt = endnoteStories.PlcfendTxt;
                EndnoteMarkerPositions = endnoteStories.MarkerPositions;
                EndnoteFormattedRuns = endnoteStories.FormattedRuns;
                EndnoteFormattedParagraphs = endnoteStories.FormattedParagraphs;
                HeaderFooterText = headerFooterStories.Text;
                PlcfHdd = headerFooterStories.PlcfHdd;
                HeaderFooterMarkerPositions = headerFooterStories.MarkerPositions;
                HeaderFooterFormattedRuns = headerFooterStories.FormattedRuns;
                HeaderFooterFormattedParagraphs = headerFooterStories.FormattedParagraphs;
                CommentText = commentStories.Text;
                PlcfandRef = commentStories.PlcfandRef;
                PlcfandTxt = commentStories.PlcfandTxt;
                CommentFormattedRuns = commentStories.FormattedRuns;
                CommentFormattedParagraphs = commentStories.FormattedParagraphs;
                PictureData = pictureData;
                HasPictures = hasPictures;
                FacingPages = facingPages;
                MirrorMargins = mirrorMargins;
                GutterAtTop = gutterAtTop;
                NoColumnBalance = noColumnBalance;
                DefaultTabStop = defaultTabStop;
                EndnotePosition = endnotePosition;
                TrackRevisions = trackRevisions;
                LockRevisionTracking = lockRevisionTracking;
                LegacyDocWritableBookmarks resolvedBookmarks = bookmarks.WithTerminalCharacterPosition(FullText.Length + 2);
                SttbfBkmk = resolvedBookmarks.SttbfBkmk;
                PlcfBkf = resolvedBookmarks.PlcfBkf;
                PlcfBkl = resolvedBookmarks.PlcfBkl;
                FontFamilies = styleSheet.FontFamilies
                    .Concat(ReadNumberingFontFamilies(numbering))
                    .Concat(formattedRuns.Select(run => run.Formatting.FontFamily))
                    .Concat(FootnoteFormattedRuns.Select(run => run.Formatting.FontFamily))
                    .Concat(HeaderFooterFormattedRuns.Select(run => run.Formatting.FontFamily))
                    .Concat(CommentFormattedRuns.Select(run => run.Formatting.FontFamily))
                    .Concat(EndnoteFormattedRuns.Select(run => run.Formatting.FontFamily))
                    .Where(fontFamily => !string.IsNullOrWhiteSpace(fontFamily))
                    .Select(fontFamily => fontFamily!)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                FontFamilyIndexes = FontFamilies
                    .Select((fontFamily, index) => new { fontFamily, index })
                    .ToDictionary(item => item.fontFamily, item => item.index, StringComparer.OrdinalIgnoreCase);
                ListTables = CreateWritableNumbering(numbering, styleSheet.StyleIndexes, FontFamilyIndexes);
                string[] revisionAuthors = CreateFormattedRuns()
                    .Select(run => run.Formatting.Revision.Author)
                    .Where(author => !string.IsNullOrWhiteSpace(author))
                    .Select(author => author!)
                    .Distinct(StringComparer.Ordinal)
                    .ToArray();
                RevisionAuthors = revisionAuthors.Length == 0
                    ? Array.Empty<string>()
                    : new[] { LegacyDocRevisionAuthorReader.UnknownAuthor }
                        .Concat(revisionAuthors.Where(author => !string.Equals(author, LegacyDocRevisionAuthorReader.UnknownAuthor, StringComparison.Ordinal)))
                        .ToArray();

                RevisionAuthorIndexes = RevisionAuthors
                    .Select((author, index) => new { author, index })
                    .ToDictionary(item => item.author, item => item.index, StringComparer.Ordinal);
                SttbfRMark = CreateRevisionAuthorTable(RevisionAuthors);
                FieldTables = new LegacyDocWritableFieldTables(Text, FormattedRuns, HeaderFooterText, HeaderFooterFormattedRuns,
                    FootnoteText, FootnoteFormattedRuns, CommentText, CommentFormattedRuns, EndnoteText, EndnoteFormattedRuns);
                ChpxPages = CreateChpxFkpPages(CreateFormattingSegments(), FontFamilyIndexes, RevisionAuthorIndexes);
                PapxPages = LegacyDocParagraphFormattingWriter.CreatePapxFkpPages(CreateParagraphSegments(), OleSectorSize);
                RequiresWord2000Format = PapxPages.Any(page => page.Any(segment =>
                    segment.Formatting.TableDepth > 1
                    || segment.Formatting.TableCellFitTexts.Any(enabled => enabled)
                    || segment.Formatting.TableCellNoWraps.Any(enabled => enabled)
                    || segment.Formatting.TableCellHideMarks.Any(enabled => enabled)));
                RmdThreading = RequiresWord2000Format ? CreateRevisionThreading(RevisionAuthors.Count) : Array.Empty<byte>();
            }

            internal string Text { get; }

            internal LegacyDocWritableFieldTables FieldTables { get; }

            internal LegacyDocWritableNumbering ListTables { get; }

            internal int FieldTablesOffsetInTableStream => AfterEndnoteDataOffsetInTableStream;

            private int AfterFieldTablesOffsetInTableStream => FieldTablesOffsetInTableStream + FieldTables.Length;

            internal string HeaderFooterText { get; }

            internal string FootnoteText { get; }

            internal string EndnoteText { get; }

            internal string CommentText { get; }

            internal string FullText => Text + FootnoteText + HeaderFooterText + CommentText + EndnoteText;

            internal bool HasNoteStories => HasFootnotes || HasEndnotes;

            // MS-DOC PlcPcd includes an extra terminal mark only when another
            // document part follows the main body. Otherwise Word displays it.
            internal bool NeedsTerminalPadding => FullText.Length > Text.Length;

            internal string StoredText => NeedsTerminalPadding ? FullText + "\r" : FullText;

            internal int PieceTableCharacterCount => FullText.Length + (NeedsTerminalPadding ? 1 : 0);

            internal byte[] PlcffndRef { get; }

            internal byte[] PlcffndTxt { get; }

            internal IReadOnlyList<int> FootnoteMarkerPositions { get; }

            internal IReadOnlyList<LegacyDocWritableRun> FootnoteFormattedRuns { get; }

            internal IReadOnlyList<LegacyDocWritableParagraph> FootnoteFormattedParagraphs { get; }

            internal byte[] PlcfendRef { get; }

            internal byte[] PlcfendTxt { get; }

            internal IReadOnlyList<int> EndnoteMarkerPositions { get; }

            internal IReadOnlyList<LegacyDocWritableRun> EndnoteFormattedRuns { get; }

            internal IReadOnlyList<LegacyDocWritableParagraph> EndnoteFormattedParagraphs { get; }

            internal IReadOnlyList<int> HeaderFooterMarkerPositions { get; }

            internal IReadOnlyList<LegacyDocWritableRun> HeaderFooterFormattedRuns { get; }

            internal IReadOnlyList<LegacyDocWritableParagraph> HeaderFooterFormattedParagraphs { get; }

            internal byte[] PlcfandRef { get; }

            internal byte[] PlcfandTxt { get; }

            internal IReadOnlyList<LegacyDocWritableRun> CommentFormattedRuns { get; }

            internal IReadOnlyList<LegacyDocWritableParagraph> CommentFormattedParagraphs { get; }

            internal byte[] PlcfHdd { get; }

            internal byte[] SttbfBkmk { get; }

            internal byte[] PlcfBkf { get; }

            internal byte[] PlcfBkl { get; }

            internal bool FacingPages { get; }

            internal bool MirrorMargins { get; }

            internal bool GutterAtTop { get; }

            internal bool NoColumnBalance { get; }

            internal ushort DefaultTabStop { get; }

            internal byte[] PictureData { get; }

            internal bool HasPictures { get; }

            internal EndnotePositionValues? EndnotePosition { get; }

            internal bool TrackRevisions { get; }

            internal bool LockRevisionTracking { get; }

            internal IReadOnlyList<LegacyDocWritableRun> FormattedRuns { get; }

            internal IReadOnlyList<LegacyDocWritableParagraph> FormattedParagraphs { get; }

            internal IReadOnlyList<LegacyDocWritableSection> Sections { get; }

            internal LegacyDocWritableStyleSheet StyleSheet { get; }

            internal IReadOnlyList<string> FontFamilies { get; }

            internal IReadOnlyDictionary<string, int> FontFamilyIndexes { get; }

            internal IReadOnlyList<string> RevisionAuthors { get; }

            internal IReadOnlyDictionary<string, int> RevisionAuthorIndexes { get; }

            internal byte[] SttbfRMark { get; }

            internal bool RequiresWord2000Format { get; }

            internal byte[] RmdThreading { get; }

            internal bool HasCharacterFormatting => true;

            internal bool HasParagraphFormatting => true;

            internal bool HasFontTable => FontFamilies.Count > 0;

            internal bool HasStyleSheet => StyleSheet.Bytes.Length > 0;

            internal bool HasRevisions => SttbfRMark.Length > 0;

            internal bool HasSectionDescriptors => Sections.Count > 0;

            internal bool HasHeaderFooterStories => HeaderFooterText.Length > 0 && PlcfHdd.Length > 0;

            internal bool HasFootnotes => FootnoteText.Length > 0 && PlcffndRef.Length > 0 && PlcffndTxt.Length > 0;

            internal bool HasEndnotes => EndnoteText.Length > 0 && PlcfendRef.Length > 0 && PlcfendTxt.Length > 0;

            internal bool HasComments => CommentText.Length > 0
                && PlcfandRef.Length > 0
                && PlcfandTxt.Length > 0;

            internal bool HasBookmarks => SttbfBkmk.Length > 0 && PlcfBkf.Length > 0 && PlcfBkl.Length > 0;

            internal int DopLength => RequiresWord2000Format ? Dop2000Length : Dop97Length;

            internal IReadOnlyList<IReadOnlyList<LegacyDocWritableSegment>> ChpxPages { get; }

            internal int ChpxPageCount => ChpxPages.Count;

            internal int ChpxPlcLength => sizeof(int) + (ChpxPageCount * sizeof(int) * 2);

            internal IReadOnlyList<IReadOnlyList<LegacyDocWritableParagraphSegment>> PapxPages { get; }

            internal int PapxPageCount => PapxPages.Count;

            internal int PapxPlcLength => sizeof(int) + (PapxPageCount * sizeof(int) * 2);

            internal int PapxPlcOffsetInTableStream => ClxLength + (HasCharacterFormatting ? ChpxPlcLength : 0);

            internal int SedPlcOffsetInTableStream => ClxLength + (HasCharacterFormatting ? ChpxPlcLength : 0) + (HasParagraphFormatting ? PapxPlcLength : 0);

            internal int SedPlcLength => 4 + (Sections.Count * (4 + SedLength));

            private int AfterSectionDataOffsetInTableStream => ClxLength + (HasCharacterFormatting ? ChpxPlcLength : 0) + (HasParagraphFormatting ? PapxPlcLength : 0) + (HasSectionDescriptors ? SedPlcLength : 0);

            internal int PlcffndRefOffsetInTableStream => AfterSectionDataOffsetInTableStream;

            internal int PlcffndTxtOffsetInTableStream => AfterSectionDataOffsetInTableStream + (HasFootnotes ? PlcffndRef.Length : 0);

            private int AfterFootnoteDataOffsetInTableStream => AfterSectionDataOffsetInTableStream + (HasFootnotes ? PlcffndRef.Length + PlcffndTxt.Length : 0);

            internal int PlcfHddOffsetInTableStream => AfterFootnoteDataOffsetInTableStream;

            private int AfterHeaderFooterDataOffsetInTableStream => AfterFootnoteDataOffsetInTableStream + (HasHeaderFooterStories ? PlcfHdd.Length : 0);

            internal int PlcfandRefOffsetInTableStream => AfterHeaderFooterDataOffsetInTableStream;

            internal int PlcfandTxtOffsetInTableStream => PlcfandRefOffsetInTableStream + (HasComments ? PlcfandRef.Length : 0);

            private int AfterCommentDataOffsetInTableStream => AfterHeaderFooterDataOffsetInTableStream
                + (HasComments ? PlcfandRef.Length + PlcfandTxt.Length : 0);

            internal int PlcfendRefOffsetInTableStream => AfterCommentDataOffsetInTableStream;

            internal int PlcfendTxtOffsetInTableStream => AfterCommentDataOffsetInTableStream + (HasEndnotes ? PlcfendRef.Length : 0);

            private int AfterEndnoteDataOffsetInTableStream => AfterCommentDataOffsetInTableStream + (HasEndnotes ? PlcfendRef.Length + PlcfendTxt.Length : 0);

            internal int DopOffsetInTableStream => AlignToEven(AfterFieldTablesOffsetInTableStream);

            internal int SttbfAssocOffsetInTableStream => DopOffsetInTableStream + DopLength;

            private int AfterDocumentOptionsOffsetInTableStream => SttbfAssocOffsetInTableStream + SttbfAssocLength;

            internal int SttbfBkmkOffsetInTableStream => HasBookmarks ? AlignToEven(AfterDocumentOptionsOffsetInTableStream) : AfterDocumentOptionsOffsetInTableStream;

            internal int PlcfBkfOffsetInTableStream => SttbfBkmkOffsetInTableStream + (HasBookmarks ? SttbfBkmk.Length : 0);

            internal int PlcfBklOffsetInTableStream => PlcfBkfOffsetInTableStream + (HasBookmarks ? PlcfBkf.Length : 0);

            private int AfterBookmarkDataOffsetInTableStream => HasBookmarks ? PlcfBklOffsetInTableStream + PlcfBkl.Length : AfterDocumentOptionsOffsetInTableStream;

            internal int SttbfRMarkOffsetInTableStream => HasRevisions ? AlignToEven(AfterBookmarkDataOffsetInTableStream) : AfterBookmarkDataOffsetInTableStream;

            private int AfterRevisionDataOffsetInTableStream => HasRevisions
                ? SttbfRMarkOffsetInTableStream + SttbfRMark.Length
                : AfterBookmarkDataOffsetInTableStream;

            internal int RmdThreadingOffsetInTableStream => RequiresWord2000Format
                ? AlignToEven(AfterRevisionDataOffsetInTableStream) : AfterRevisionDataOffsetInTableStream;

            private int AfterRevisionThreadingOffsetInTableStream => RmdThreadingOffsetInTableStream + RmdThreading.Length;

            internal int StyleSheetOffsetInTableStream => HasStyleSheet ? AlignToEven(AfterRevisionThreadingOffsetInTableStream) : AfterRevisionThreadingOffsetInTableStream;

            private int AfterStyleSheetOffsetInTableStream => HasStyleSheet
                ? StyleSheetOffsetInTableStream + StyleSheet.Bytes.Length
                : AfterRevisionThreadingOffsetInTableStream;

            internal int ListDefinitionsOffsetInTableStream => AlignToEven(AfterStyleSheetOffsetInTableStream);

            internal int ListInstancesOffsetInTableStream => ListDefinitionsOffsetInTableStream + ListTables.Definitions.Length;

            internal int FontTableOffsetInTableStream => ListTables.Definitions.Length == 0 ? AfterStyleSheetOffsetInTableStream
                : ListInstancesOffsetInTableStream + ListTables.Instances.Length;

            internal IReadOnlyList<LegacyDocWritableSegment> CreateFormattingSegments() {
                var segments = new List<LegacyDocWritableSegment>();
                int character = 0;
                foreach (LegacyDocWritableRun run in CreateFormattedRuns().OrderBy(item => item.StartCharacter)) {
                    if (run.StartCharacter > character) {
                        AddSegment(segments, character, run.StartCharacter - character, LegacyDocWritableFormatting.Plain);
                    }

                    AddSegment(segments, run.StartCharacter, run.Length, run.Formatting, run.PictureDataOffset);
                    character = run.EndCharacter;
                }

                if (character < PieceTableCharacterCount) {
                    AddSegment(segments, character, PieceTableCharacterCount - character, LegacyDocWritableFormatting.Plain);
                }

                return segments;
            }

            private IReadOnlyList<LegacyDocWritableRun> CreateFormattedRuns() {
                if (FootnoteMarkerPositions.Count == 0
                    && FootnoteFormattedRuns.Count == 0
                    && HeaderFooterMarkerPositions.Count == 0
                    && HeaderFooterFormattedRuns.Count == 0
                    && CommentFormattedRuns.Count == 0
                    && EndnoteMarkerPositions.Count == 0
                    && EndnoteFormattedRuns.Count == 0) {
                    return FormattedRuns;
                }

                var runs = new List<LegacyDocWritableRun>(
                    FormattedRuns.Count
                    + FootnoteMarkerPositions.Count
                    + FootnoteFormattedRuns.Count
                    + HeaderFooterMarkerPositions.Count
                    + HeaderFooterFormattedRuns.Count
                    + CommentFormattedRuns.Count
                    + EndnoteMarkerPositions.Count
                    + EndnoteFormattedRuns.Count);
                runs.AddRange(FormattedRuns);
                int footnoteStartCharacter = Text.Length;
                foreach (LegacyDocWritableRun run in FootnoteFormattedRuns) {
                    runs.Add(new LegacyDocWritableRun(footnoteStartCharacter + run.StartCharacter, run.Length, run.Formatting, run.PictureDataOffset));
                }

                foreach (int markerPosition in FootnoteMarkerPositions) {
                    runs.Add(new LegacyDocWritableRun(footnoteStartCharacter + markerPosition, 1, LegacyDocWritableFormatting.SpecialCharacter));
                }

                int headerFooterStartCharacter = Text.Length + FootnoteText.Length;
                foreach (LegacyDocWritableRun run in HeaderFooterFormattedRuns) {
                    runs.Add(new LegacyDocWritableRun(headerFooterStartCharacter + run.StartCharacter, run.Length, run.Formatting, run.PictureDataOffset));
                }

                foreach (int markerPosition in HeaderFooterMarkerPositions) {
                    runs.Add(new LegacyDocWritableRun(headerFooterStartCharacter + markerPosition, 1, LegacyDocWritableFormatting.SpecialCharacter));
                }

                int commentStartCharacter = Text.Length + FootnoteText.Length + HeaderFooterText.Length;
                foreach (LegacyDocWritableRun run in CommentFormattedRuns) {
                    runs.Add(new LegacyDocWritableRun(
                        commentStartCharacter + run.StartCharacter,
                        run.Length,
                        run.Formatting,
                        run.PictureDataOffset));
                }

                int endnoteStartCharacter = commentStartCharacter + CommentText.Length;
                foreach (LegacyDocWritableRun run in EndnoteFormattedRuns) {
                    runs.Add(new LegacyDocWritableRun(endnoteStartCharacter + run.StartCharacter, run.Length, run.Formatting, run.PictureDataOffset));
                }

                foreach (int markerPosition in EndnoteMarkerPositions) {
                    runs.Add(new LegacyDocWritableRun(endnoteStartCharacter + markerPosition, 1, LegacyDocWritableFormatting.SpecialCharacter));
                }

                return runs;
            }

            private static void AddSegment(
                List<LegacyDocWritableSegment> segments,
                int startCharacter,
                int length,
                LegacyDocWritableFormatting formatting,
                int? pictureDataOffset = null) {
                if (length <= 0) {
                    return;
                }

                if (segments.Count > 0) {
                    LegacyDocWritableSegment previous = segments[segments.Count - 1];
                    if (previous.EndCharacter == startCharacter
                        && previous.Formatting.HasSameEncoding(formatting)
                        && previous.PictureDataOffset == pictureDataOffset) {
                        segments[segments.Count - 1] = previous.Extend(length);
                        return;
                    }
                }

                segments.Add(new LegacyDocWritableSegment(startCharacter, length, formatting, pictureDataOffset));
            }

            internal IReadOnlyList<LegacyDocWritableParagraphSegment> CreateParagraphSegments() {
                // Word finds paragraph boundaries in PAPX, including unformatted paragraphs.
                return CreateStoryParagraphSegments();
            }

            private IReadOnlyList<LegacyDocWritableParagraphSegment> CreateStoryParagraphSegments() {
                var segments = new List<LegacyDocWritableParagraphSegment>();
                AddBodyParagraphSegments(segments);
                AddStoryParagraphSegments(
                    segments,
                    FootnoteText,
                    Text.Length,
                    CreateNoteParagraphFormatter(FootnoteFormattedParagraphs, Text.Length));
                AddStoryParagraphSegments(segments, HeaderFooterText, Text.Length + FootnoteText.Length, CreateHeaderFooterParagraphFormatter());
                int commentStartCharacter = Text.Length + FootnoteText.Length + HeaderFooterText.Length;
                AddStoryParagraphSegments(
                    segments,
                    CommentText,
                    commentStartCharacter,
                    CreateNoteParagraphFormatter(CommentFormattedParagraphs, commentStartCharacter));
                int endnoteStartCharacter = commentStartCharacter + CommentText.Length;
                AddStoryParagraphSegments(
                    segments,
                    EndnoteText,
                    endnoteStartCharacter,
                    CreateNoteParagraphFormatter(EndnoteFormattedParagraphs, endnoteStartCharacter));
                AddRawParagraphSegment(segments, FullText.Length, PieceTableCharacterCount - FullText.Length, PlainParagraphPapx);
                return segments;
            }

            private static Func<LegacyDocWritableParagraphRange, object> CreateNoteParagraphFormatter(IReadOnlyList<LegacyDocWritableParagraph> storyFormattedParagraphs, int storyStartCharacter) {
                if (storyFormattedParagraphs.Count == 0) {
                    return CreatePlainNoteParagraphFormatter();
                }

                LegacyDocWritableParagraph[] formattedParagraphs = storyFormattedParagraphs
                    .OrderBy(item => item.StartCharacter)
                    .ToArray();
                int formattedIndex = 0;
                return paragraph => {
                    while (formattedIndex < formattedParagraphs.Length
                        && storyStartCharacter + formattedParagraphs[formattedIndex].EndCharacter <= paragraph.Start) {
                        formattedIndex++;
                    }

                    if (formattedIndex < formattedParagraphs.Length
                        && storyStartCharacter + formattedParagraphs[formattedIndex].StartCharacter == paragraph.Start
                        && formattedParagraphs[formattedIndex].Length == paragraph.Length) {
                        return formattedParagraphs[formattedIndex].Formatting;
                    }

                    return CreatePlainNoteParagraphFormat(paragraph);
                };
            }

            private static Func<LegacyDocWritableParagraphRange, object> CreatePlainNoteParagraphFormatter() {
                return CreatePlainNoteParagraphFormat;
            }

            private static object CreatePlainNoteParagraphFormat(LegacyDocWritableParagraphRange paragraph) {
                return paragraph.Length > 0 && paragraph.Text[0] == LegacyDocFootnoteReader.FootnoteReferenceCharacter
                    ? FootnoteTextParagraphPapx
                    : PlainParagraphPapx;
            }

            private Func<LegacyDocWritableParagraphRange, object> CreateHeaderFooterParagraphFormatter() {
                if (HeaderFooterFormattedParagraphs.Count == 0) {
                    return _ => PlainParagraphPapx;
                }

                LegacyDocWritableParagraph[] formattedParagraphs = HeaderFooterFormattedParagraphs
                    .OrderBy(item => item.StartCharacter)
                    .ToArray();
                int headerFooterStartCharacter = Text.Length + FootnoteText.Length;
                int formattedIndex = 0;
                return paragraph => {
                    while (formattedIndex < formattedParagraphs.Length
                        && headerFooterStartCharacter + formattedParagraphs[formattedIndex].EndCharacter <= paragraph.Start) {
                        formattedIndex++;
                    }

                    if (formattedIndex < formattedParagraphs.Length
                        && headerFooterStartCharacter + formattedParagraphs[formattedIndex].StartCharacter == paragraph.Start
                        && formattedParagraphs[formattedIndex].Length == paragraph.Length) {
                        return formattedParagraphs[formattedIndex].Formatting;
                    }

                    return PlainParagraphPapx;
                };
            }

            private void AddBodyParagraphSegments(List<LegacyDocWritableParagraphSegment> segments) {
                var formattedParagraphs = FormattedParagraphs
                    .OrderBy(item => item.StartCharacter)
                    .ToArray();
                int formattedIndex = 0;
                AddStoryParagraphSegments(
                    segments,
                    Text,
                    0,
                    paragraph => {
                        while (formattedIndex < formattedParagraphs.Length
                            && formattedParagraphs[formattedIndex].EndCharacter <= paragraph.Start) {
                            formattedIndex++;
                        }

                        if (formattedIndex < formattedParagraphs.Length
                            && formattedParagraphs[formattedIndex].StartCharacter == paragraph.Start
                            && formattedParagraphs[formattedIndex].Length == paragraph.Length) {
                            return formattedParagraphs[formattedIndex].Formatting;
                        }

                        return PlainParagraphPapx;
                    }, new HashSet<int>(Sections.Select(section => section.EndCharacter)));
            }

            private static void AddStoryParagraphSegments(
                List<LegacyDocWritableParagraphSegment> segments,
                string story,
                int storyStart,
                Func<LegacyDocWritableParagraphRange, object> selectParagraphFormat,
                HashSet<int>? sectionEndCharacters = null) {
                int paragraphStart = 0;
                for (int index = 0; index < story.Length; index++) {
                    bool isSectionMark = story[index] == LegacyDocSpecialCharacters.PageBreak &&
                        sectionEndCharacters?.Contains(storyStart + index + 1) == true;
                    if (story[index] != '\r' && story[index] != '\a' && !isSectionMark) {
                        continue;
                    }

                    AddStoryParagraphSegment(segments, story, storyStart, paragraphStart, index + 1, selectParagraphFormat);
                    paragraphStart = index + 1;
                }

                if (paragraphStart < story.Length) {
                    AddStoryParagraphSegment(segments, story, storyStart, paragraphStart, story.Length, selectParagraphFormat);
                }
            }

            private static void AddStoryParagraphSegment(
                List<LegacyDocWritableParagraphSegment> segments,
                string story,
                int storyStart,
                int paragraphStart,
                int paragraphEnd,
                Func<LegacyDocWritableParagraphRange, object> selectParagraphFormat) {
                int length = paragraphEnd - paragraphStart;
                if (length <= 0) {
                    return;
                }

                var paragraph = new LegacyDocWritableParagraphRange(storyStart + paragraphStart, length, story.Substring(paragraphStart, length));
                object paragraphFormat = selectParagraphFormat(paragraph);
                if (paragraphFormat is LegacyDocWritableParagraphFormatting formatting) {
                    AddParagraphSegment(segments, paragraph.Start, paragraph.Length, formatting);
                } else if (paragraphFormat is byte[] papxOverride) {
                    AddRawParagraphSegment(segments, paragraph.Start, paragraph.Length, papxOverride);
                } else {
                    throw new InvalidOperationException("The generated DOC paragraph segment formatter returned an unsupported value.");
                }
            }

            private static void AddParagraphSegment(
                List<LegacyDocWritableParagraphSegment> segments,
                int startCharacter,
                int length,
                LegacyDocWritableParagraphFormatting formatting) {
                if (length <= 0) {
                    return;
                }

                // PAPX boundaries identify paragraphs to Word, even when their properties
                // are identical. Coalescing them turns paragraph marks into inline breaks.
                segments.Add(new LegacyDocWritableParagraphSegment(startCharacter, length, formatting));
            }

            private static void AddRawParagraphSegment(
                List<LegacyDocWritableParagraphSegment> segments,
                int startCharacter,
                int length,
                byte[] papxOverride) {
                if (length <= 0) {
                    return;
                }

                segments.Add(new LegacyDocWritableParagraphSegment(startCharacter, length, papxOverride));
            }
        }

        private readonly struct LegacyDocWritableParagraphRange {
            internal LegacyDocWritableParagraphRange(int start, int length, string text) {
                Start = start;
                Length = length;
                Text = text;
            }

            internal int Start { get; }

            internal int Length { get; }

            internal string Text { get; }

            internal char this[int index] => Text[index];
        }

    }
}
