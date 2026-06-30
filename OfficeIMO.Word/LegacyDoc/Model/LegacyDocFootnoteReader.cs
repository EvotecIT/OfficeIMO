namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocFootnoteReader {
        internal const char FootnoteReferenceCharacter = '\u0002';

        internal static bool HasReadableFootnoteTables(byte[] tableStream, LegacyDocFib fib) {
            return fib.CcpFtn == 0
                || TryReadNoteTables(
                    tableStream,
                    fib.CcpFtn,
                    fib.FcPlcffndRef,
                    fib.LcbPlcffndRef,
                    fib.FcPlcffndTxt,
                    fib.LcbPlcffndTxt,
                    "footnote",
                    out _,
                    out _,
                    out _);
        }

        internal static bool HasReadableEndnoteTables(byte[] tableStream, LegacyDocFib fib) {
            return fib.CcpEdn == 0
                || TryReadNoteTables(
                    tableStream,
                    fib.CcpEdn,
                    fib.FcPlcfendRef,
                    fib.LcbPlcfendRef,
                    fib.FcPlcfendTxt,
                    fib.LcbPlcfendTxt,
                    "endnote",
                    out _,
                    out _,
                    out _);
        }

        internal static IReadOnlyList<LegacyDocFootnote> Read(
            byte[] tableStream,
            LegacyDocTextContent textContent,
            LegacyDocFib fib,
            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges,
            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphFormattingRanges,
            LegacyDocBookmarkProjectionTracker bookmarkProjection,
            out string? warning) {
            warning = null;
            if (fib.CcpFtn == 0) {
                return Array.Empty<LegacyDocFootnote>();
            }

            if (!TryReadNoteTables(
                tableStream,
                fib.CcpFtn,
                fib.FcPlcffndRef,
                fib.LcbPlcffndRef,
                fib.FcPlcffndTxt,
                fib.LcbPlcffndTxt,
                "footnote",
                out int[] referencePositions,
                out int[] textPositions,
                out warning)) {
                return Array.Empty<LegacyDocFootnote>();
            }

            int footnoteBaseCharacterPosition = fib.CcpText;
            int footnoteCount = referencePositions.Length;
            var footnotes = new List<LegacyDocFootnote>(footnoteCount);
            for (int index = 0; index < footnoteCount; index++) {
                int startCharacter = textPositions[index];
                int endCharacter = textPositions[index + 1];
                if (endCharacter <= startCharacter) {
                    continue;
                }

                IReadOnlyList<LegacyDocNoteParagraph> paragraphs = BuildStoryParagraphs(
                    textContent.AllCharacters,
                    footnoteBaseCharacterPosition + startCharacter,
                    footnoteBaseCharacterPosition + endCharacter,
                    formattingRanges,
                    paragraphFormattingRanges,
                    bookmarkProjection);
                if (paragraphs.Count == 0) {
                    continue;
                }

                footnotes.Add(new LegacyDocFootnote(referencePositions[index], paragraphs));
            }

            return footnotes;
        }

        internal static IReadOnlyList<LegacyDocEndnote> ReadEndnotes(
            byte[] tableStream,
            LegacyDocTextContent textContent,
            LegacyDocFib fib,
            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges,
            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphFormattingRanges,
            LegacyDocBookmarkProjectionTracker bookmarkProjection,
            out string? warning) {
            warning = null;
            if (fib.CcpEdn == 0) {
                return Array.Empty<LegacyDocEndnote>();
            }

            if (!TryReadNoteTables(
                tableStream,
                fib.CcpEdn,
                fib.FcPlcfendRef,
                fib.LcbPlcfendRef,
                fib.FcPlcfendTxt,
                fib.LcbPlcfendTxt,
                "endnote",
                out int[] referencePositions,
                out int[] textPositions,
                out warning)) {
                return Array.Empty<LegacyDocEndnote>();
            }

            int endnoteBaseCharacterPosition = fib.CcpText + fib.CcpFtn + fib.CcpHdd + fib.CcpAtn;
            int endnoteCount = referencePositions.Length;
            var endnotes = new List<LegacyDocEndnote>(endnoteCount);
            for (int index = 0; index < endnoteCount; index++) {
                int startCharacter = textPositions[index];
                int endCharacter = textPositions[index + 1];
                if (endCharacter <= startCharacter) {
                    continue;
                }

                IReadOnlyList<LegacyDocNoteParagraph> paragraphs = BuildStoryParagraphs(
                    textContent.AllCharacters,
                    endnoteBaseCharacterPosition + startCharacter,
                    endnoteBaseCharacterPosition + endCharacter,
                    formattingRanges,
                    paragraphFormattingRanges,
                    bookmarkProjection);
                if (paragraphs.Count == 0) {
                    continue;
                }

                endnotes.Add(new LegacyDocEndnote(referencePositions[index], paragraphs));
            }

            return endnotes;
        }

        private static bool TryReadNoteTables(
            byte[] tableStream,
            int storyCharacterCount,
            int fcReference,
            int lcbReference,
            int fcText,
            int lcbText,
            string noteKind,
            out int[] referencePositions,
            out int[] textPositions,
            out string? warning) {
            referencePositions = Array.Empty<int>();
            textPositions = Array.Empty<int>();
            warning = null;

            if (storyCharacterCount == 0) {
                return true;
            }

            if (!TryReadNoteReferencePositions(tableStream, fcReference, lcbReference, noteKind, out referencePositions, out warning)) {
                return false;
            }

            if (!TryReadNoteTextPositions(tableStream, fcText, lcbText, noteKind, out textPositions, out warning)) {
                return false;
            }

            if (referencePositions.Length == 0 || textPositions.Length < referencePositions.Length + 1) {
                warning = $"The {noteKind} reference and text PLCs do not contain matching simple {noteKind} ranges.";
                referencePositions = Array.Empty<int>();
                textPositions = Array.Empty<int>();
                return false;
            }

            textPositions = textPositions.Take(referencePositions.Length + 1).ToArray();
            int previousTextPosition = -1;
            for (int index = 0; index < textPositions.Length; index++) {
                int position = textPositions[index];
                if (position < previousTextPosition || position < 0 || position > storyCharacterCount) {
                    warning = $"The {noteKind} text PLC contains a non-monotonic or out-of-range character position.";
                    referencePositions = Array.Empty<int>();
                    textPositions = Array.Empty<int>();
                    return false;
                }

                previousTextPosition = position;
            }

            return true;
        }

        private static bool TryReadNoteReferencePositions(byte[] tableStream, int fcReference, int lcbReference, string noteKind, out int[] positions, out string? warning) {
            positions = Array.Empty<int>();
            warning = null;
            if (lcbReference == 0) {
                warning = $"The FIB reports {noteKind} story text without a {noteKind} reference PLC.";
                return false;
            }

            if (fcReference < 0
                || lcbReference < 4
                || fcReference + lcbReference > tableStream.Length
                || (lcbReference - 4) % 6 != 0) {
                warning = $"The FIB points outside the selected table stream for the {noteKind} reference PLC.";
                return false;
            }

            int noteCount = (lcbReference - 4) / 6;
            var cps = new int[noteCount + 1];
            for (int index = 0; index < cps.Length; index++) {
                cps[index] = LegacyDocFib.ReadInt32(tableStream, fcReference + (index * 4));
            }

            positions = cps.Take(noteCount).ToArray();
            return true;
        }

        private static bool TryReadNoteTextPositions(byte[] tableStream, int fcText, int lcbText, string noteKind, out int[] positions, out string? warning) {
            positions = Array.Empty<int>();
            warning = null;
            if (lcbText == 0) {
                warning = $"The FIB reports {noteKind} story text without a {noteKind} text PLC.";
                return false;
            }

            if (fcText < 0
                || lcbText < 8
                || fcText + lcbText > tableStream.Length
                || lcbText % 4 != 0) {
                warning = $"The FIB points outside the selected table stream for the {noteKind} text PLC.";
                return false;
            }

            int positionCount = lcbText / 4;
            positions = new int[positionCount];
            for (int index = 0; index < positionCount; index++) {
                positions[index] = LegacyDocFib.ReadInt32(tableStream, fcText + (index * 4));
            }

            return true;
        }

        private static IReadOnlyList<LegacyDocNoteParagraph> BuildStoryParagraphs(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            int startCharacter,
            int endCharacter,
            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges,
            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphFormattingRanges,
            LegacyDocBookmarkProjectionTracker bookmarkProjection) {
            if (endCharacter <= startCharacter) {
                return Array.Empty<LegacyDocNoteParagraph>();
            }

            var paragraphs = new List<LegacyDocNoteParagraph>();
            var currentRuns = new List<LegacyDocTextRun>();
            var runText = new System.Text.StringBuilder(endCharacter - startCharacter);
            var runCharacterPositions = new List<int>();
            LegacyDocCharacterFormat currentFormat = LegacyDocCharacterFormat.Default;
            LegacyDocHyperlinkTarget currentHyperlinkTarget = default;
            bool hasCurrentRun = false;
            bool isFirstParagraph = true;
            bool atParagraphStart = true;
            bool skipOptionalReferenceSpace = false;
            int currentParagraphStartCharacter = startCharacter;

            LegacyDocTextCharacter[] storyCharacters = characters
                .Where(character => character.CharacterPosition >= startCharacter && character.CharacterPosition < endCharacter)
                .ToArray();

            for (int index = 0; index < storyCharacters.Length; index++) {
                LegacyDocTextCharacter character = storyCharacters[index];

                if (LegacyDocField.TryReadHyperlink(
                    storyCharacters,
                    index,
                    out LegacyDocHyperlinkTarget hyperlinkTarget,
                    out int resultStartIndex,
                    out int resultEndIndex,
                    out int fieldEndIndex)) {
                    for (int resultIndex = resultStartIndex; resultIndex < resultEndIndex; resultIndex++) {
                        LegacyDocTextCharacter resultCharacter = storyCharacters[resultIndex];
                        AppendRunCharacter(
                            resultCharacter.Character,
                            GetFormatForFileOffset(formattingRanges, resultCharacter.FileOffset),
                            resultCharacter.CharacterPosition,
                            hyperlinkTarget);
                    }

                    index = fieldEndIndex;
                    atParagraphStart = false;
                    skipOptionalReferenceSpace = false;
                    continue;
                }

                char normalized = character.Character == '\a' ? '\r' : character.Character;
                if (isFirstParagraph && atParagraphStart && normalized == FootnoteReferenceCharacter) {
                    skipOptionalReferenceSpace = true;
                    continue;
                }

                if (skipOptionalReferenceSpace) {
                    skipOptionalReferenceSpace = false;
                    if (normalized == ' ') {
                        continue;
                    }
                }

                atParagraphStart = false;
                if (normalized == '\r') {
                    AddCurrentParagraph(GetParagraphFormatForFileOffset(paragraphFormattingRanges, character.FileOffset), character.CharacterPosition);
                    isFirstParagraph = false;
                    atParagraphStart = true;
                    skipOptionalReferenceSpace = false;
                    currentParagraphStartCharacter = character.CharacterPosition + 1;
                    continue;
                }

                if (char.IsControl(normalized) && normalized != '\t') {
                    continue;
                }

                AppendRunCharacter(
                    normalized,
                    GetFormatForFileOffset(formattingRanges, character.FileOffset),
                    character.CharacterPosition,
                    default);
            }

            AddCurrentParagraph(LegacyDocParagraphFormat.Default, endCharacter);
            return paragraphs;

            void AppendRunCharacter(char value, LegacyDocCharacterFormat format, int characterPosition, LegacyDocHyperlinkTarget hyperlinkTarget) {
                if (!hasCurrentRun
                    || !format.Equals(currentFormat)
                    || hyperlinkTarget != currentHyperlinkTarget) {
                    FlushRun();
                    currentFormat = format;
                    currentHyperlinkTarget = hyperlinkTarget;
                    hasCurrentRun = true;
                }

                runText.Append(value);
                runCharacterPositions.Add(characterPosition);
            }

            void AddCurrentParagraph(LegacyDocParagraphFormat format, int paragraphEndCharacter) {
                FlushRun();
                if (currentRuns.Count > 0) {
                    paragraphs.Add(new LegacyDocNoteParagraph(
                        currentRuns.ToArray(),
                        format,
                        currentParagraphStartCharacter,
                        paragraphEndCharacter,
                        bookmarkProjection.ExtractProjectedParagraphBookmarks(currentParagraphStartCharacter, paragraphEndCharacter)));
                    currentRuns.Clear();
                }

                hasCurrentRun = false;
                currentHyperlinkTarget = default;
            }

            void FlushRun() {
                if (runText.Length == 0) {
                    return;
                }

                currentRuns.Add(new LegacyDocTextRun(
                    runText.ToString(),
                    currentFormat.Bold,
                    currentFormat.Italic,
                    currentFormat.Strike,
                    currentFormat.DoubleStrike,
                    currentFormat.Outline,
                    currentFormat.Shadow,
                    currentFormat.Emboss,
                    currentFormat.Imprint,
                    currentFormat.Hidden,
                    currentFormat.NoProof,
                    currentFormat.Caps,
                    currentFormat.VerticalPosition,
                    currentFormat.Underline,
                    currentFormat.Highlight,
                    currentFormat.FontSizeHalfPoints,
                    currentFormat.ColorHex,
                    currentFormat.FontFamily,
                    runCharacterPositions,
                    currentHyperlinkTarget.Uri,
                    currentHyperlinkTarget.Anchor));
                runText.Clear();
                runCharacterPositions.Clear();
            }
        }

        private static LegacyDocCharacterFormat GetFormatForFileOffset(IReadOnlyList<LegacyDocCharacterFormatRange> ranges, int fileOffset) {
            for (int i = 0; i < ranges.Count; i++) {
                if (ranges[i].Contains(fileOffset)) {
                    return ranges[i].Format;
                }
            }

            return LegacyDocCharacterFormat.Default;
        }

        private static LegacyDocParagraphFormat GetParagraphFormatForFileOffset(IReadOnlyList<LegacyDocParagraphFormatRange> ranges, int fileOffset) {
            for (int i = 0; i < ranges.Count; i++) {
                if (ranges[i].Contains(fileOffset)) {
                    return ranges[i].Format;
                }
            }

            return LegacyDocParagraphFormat.Default;
        }
    }
}
