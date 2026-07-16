using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocCommentReader {
        internal const char CommentReferenceCharacter = '\u0005';
        private const int AtrdPre10Size = 30;

        internal static bool HasReadableCommentTables(byte[] tableStream, LegacyDocFib fib) {
            return fib.CcpAtn == 0
                || TryReadCommentTables(
                    tableStream,
                    fib,
                    out _,
                    out _,
                    out _,
                    out _);
        }

        internal static IReadOnlyList<LegacyDocComment> Read(
            byte[] tableStream,
            LegacyDocTextContent textContent,
            LegacyDocFib fib,
            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges,
            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphFormattingRanges,
            LegacyDocBookmarkProjectionTracker bookmarkProjection,
            IReadOnlyDictionary<int, LegacyDocPicture> picturesByCharacterPosition,
            out string? warning) {
            warning = null;
            if (fib.CcpAtn == 0) {
                return Array.Empty<LegacyDocComment>();
            }

            if (!TryReadCommentTables(
                tableStream,
                fib,
                out int[] referencePositions,
                out int[] textPositions,
                out string[] initials,
                out warning)) {
                return Array.Empty<LegacyDocComment>();
            }

            int commentBaseCharacterPosition = fib.CcpText + fib.CcpFtn + fib.CcpHdd;
            var comments = new List<LegacyDocComment>(referencePositions.Length);
            for (int index = 0; index < referencePositions.Length; index++) {
                int startCharacter = textPositions[index];
                int endCharacter = textPositions[index + 1];
                if (endCharacter <= startCharacter) {
                    continue;
                }

                IReadOnlyList<LegacyDocNoteParagraph> paragraphs = LegacyDocFootnoteReader.BuildStoryParagraphs(
                    textContent.AllCharacters,
                    commentBaseCharacterPosition + startCharacter,
                    commentBaseCharacterPosition + endCharacter,
                    formattingRanges,
                    paragraphFormattingRanges,
                    bookmarkProjection,
                    picturesByCharacterPosition);
                if (paragraphs.Count == 0) {
                    continue;
                }

                comments.Add(new LegacyDocComment(
                    referencePositions[index],
                    paragraphs,
                    author: "Legacy DOC",
                    initials: initials[index]));
            }

            return comments;
        }

        private static bool TryReadCommentTables(
            byte[] tableStream,
            LegacyDocFib fib,
            out int[] referencePositions,
            out int[] textPositions,
            out string[] initials,
            out string? warning) {
            referencePositions = Array.Empty<int>();
            textPositions = Array.Empty<int>();
            initials = Array.Empty<string>();
            warning = null;

            if (fib.CcpAtn == 0) {
                return true;
            }

            if (!TryReadCommentReferencePositions(tableStream, fib, out referencePositions, out initials, out warning)) {
                return false;
            }

            if (!TryReadCommentTextPositions(tableStream, fib, out textPositions, out warning)) {
                return false;
            }

            if (referencePositions.Length == 0 || textPositions.Length < referencePositions.Length + 1) {
                warning = "The comment reference and text PLCs do not contain matching simple comment ranges.";
                referencePositions = Array.Empty<int>();
                textPositions = Array.Empty<int>();
                initials = Array.Empty<string>();
                return false;
            }

            textPositions = textPositions.Take(referencePositions.Length + 1).ToArray();
            int previousTextPosition = -1;
            for (int index = 0; index < textPositions.Length; index++) {
                int position = textPositions[index];
                if (position < previousTextPosition || position < 0 || position > fib.CcpAtn) {
                    warning = "The comment text PLC contains a non-monotonic or out-of-range character position.";
                    referencePositions = Array.Empty<int>();
                    textPositions = Array.Empty<int>();
                    initials = Array.Empty<string>();
                    return false;
                }

                previousTextPosition = position;
            }

            return true;
        }

        private static bool TryReadCommentReferencePositions(byte[] tableStream, LegacyDocFib fib, out int[] positions, out string[] initials, out string? warning) {
            positions = Array.Empty<int>();
            initials = Array.Empty<string>();
            warning = null;
            if (fib.LcbPlcfandRef == 0) {
                warning = "The FIB reports comment story text without a comment reference PLC.";
                return false;
            }

            if (fib.FcPlcfandRef < 0
                || fib.LcbPlcfandRef < 4 + AtrdPre10Size
                || fib.FcPlcfandRef + fib.LcbPlcfandRef > tableStream.Length
                || (fib.LcbPlcfandRef - 4) % (4 + AtrdPre10Size) != 0) {
                warning = "The FIB points outside the selected table stream for the comment reference PLC.";
                return false;
            }

            int commentCount = (fib.LcbPlcfandRef - 4) / (4 + AtrdPre10Size);
            var cps = new int[commentCount + 1];
            for (int index = 0; index < cps.Length; index++) {
                cps[index] = LegacyDocFib.ReadInt32(tableStream, fib.FcPlcfandRef + (index * 4));
            }

            var parsedInitials = new string[commentCount];
            int atrdOffset = fib.FcPlcfandRef + ((commentCount + 1) * 4);
            for (int index = 0; index < commentCount; index++) {
                parsedInitials[index] = ReadLpxCharBuffer9(tableStream, atrdOffset + (index * AtrdPre10Size));
            }

            positions = cps.Take(commentCount).ToArray();
            initials = parsedInitials;
            return true;
        }

        private static bool TryReadCommentTextPositions(byte[] tableStream, LegacyDocFib fib, out int[] positions, out string? warning) {
            positions = Array.Empty<int>();
            warning = null;
            if (fib.LcbPlcfandTxt == 0) {
                warning = "The FIB reports comment story text without a comment text PLC.";
                return false;
            }

            if (fib.FcPlcfandTxt < 0
                || fib.LcbPlcfandTxt < 8
                || fib.FcPlcfandTxt + fib.LcbPlcfandTxt > tableStream.Length
                || fib.LcbPlcfandTxt % 4 != 0) {
                warning = "The FIB points outside the selected table stream for the comment text PLC.";
                return false;
            }

            int positionCount = fib.LcbPlcfandTxt / 4;
            positions = new int[positionCount];
            for (int index = 0; index < positionCount; index++) {
                positions[index] = LegacyDocFib.ReadInt32(tableStream, fib.FcPlcfandTxt + (index * 4));
            }

            return true;
        }

        private static string ReadLpxCharBuffer9(byte[] bytes, int offset) {
            if (offset < 0 || offset + 20 > bytes.Length) {
                return string.Empty;
            }

            int characterCount = LegacyDocFib.ReadUInt16(bytes, offset);
            if (characterCount <= 0 || characterCount > 9) {
                return string.Empty;
            }

            var builder = new StringBuilder(characterCount);
            for (int index = 0; index < characterCount; index++) {
                char value = (char)LegacyDocFib.ReadUInt16(bytes, offset + 2 + (index * 2));
                if (value == '\0') {
                    break;
                }

                builder.Append(value);
            }

            return builder.ToString();
        }
    }
}
