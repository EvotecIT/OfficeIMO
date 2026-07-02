using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocBookmarkReader {
        internal static IReadOnlyList<LegacyDocBookmark> Read(byte[] tableStream, LegacyDocFib fib, out string? warning) {
            warning = null;
            if (fib.LcbSttbfBkmk == 0 && fib.LcbPlcfBkf == 0 && fib.LcbPlcfBkl == 0) {
                return Array.Empty<LegacyDocBookmark>();
            }

            if (fib.LcbSttbfBkmk == 0 || fib.LcbPlcfBkf == 0 || fib.LcbPlcfBkl == 0) {
                warning = "The DOC bookmark table is incomplete.";
                return Array.Empty<LegacyDocBookmark>();
            }

            if (!TryReadBookmarkNames(tableStream, fib.FcSttbfBkmk, fib.LcbSttbfBkmk, out IReadOnlyList<string> names, out warning)) {
                return Array.Empty<LegacyDocBookmark>();
            }

            int count = names.Count;
            int expectedBkfLength = ((count + 1) * 4) + (count * 4);
            int expectedBklLength = (count + 1) * 4;
            if (fib.LcbPlcfBkf != expectedBkfLength || fib.LcbPlcfBkl != expectedBklLength) {
                warning = "The DOC bookmark start/end PLC tables do not match the bookmark name table.";
                return Array.Empty<LegacyDocBookmark>();
            }

            if (!IsRangeWithin(tableStream, fib.FcPlcfBkf, fib.LcbPlcfBkf)
                || !IsRangeWithin(tableStream, fib.FcPlcfBkl, fib.LcbPlcfBkl)) {
                warning = "The DOC bookmark PLC tables extend beyond the table stream.";
                return Array.Empty<LegacyDocBookmark>();
            }

            var endCharacters = new int[count];
            for (int index = 0; index < count; index++) {
                endCharacters[index] = LegacyDocFib.ReadInt32(tableStream, fib.FcPlcfBkl + (index * 4));
            }

            var bookmarks = new List<LegacyDocBookmark>(count);
            int fbkfOffset = fib.FcPlcfBkf + ((count + 1) * 4);
            for (int index = 0; index < count; index++) {
                int startCharacter = LegacyDocFib.ReadInt32(tableStream, fib.FcPlcfBkf + (index * 4));
                int endIndex = LegacyDocFib.ReadUInt16(tableStream, fbkfOffset + (index * 4));
                if (endIndex < 0 || endIndex >= count) {
                    warning = "The DOC bookmark start PLC contains an invalid end bookmark index.";
                    return Array.Empty<LegacyDocBookmark>();
                }

                int endCharacter = endCharacters[endIndex];
                if (startCharacter < 0 || endCharacter < startCharacter) {
                    warning = "The DOC bookmark PLC contains an invalid character range.";
                    return Array.Empty<LegacyDocBookmark>();
                }

                bookmarks.Add(new LegacyDocBookmark(names[index], startCharacter, endCharacter, checked(index + 1).ToString()));
            }

            return bookmarks;
        }

        private static bool TryReadBookmarkNames(byte[] tableStream, int offset, int length, out IReadOnlyList<string> names, out string? warning) {
            names = Array.Empty<string>();
            warning = null;
            if (!IsRangeWithin(tableStream, offset, length) || length < 6) {
                warning = "The DOC bookmark name table extends beyond the table stream.";
                return false;
            }

            ushort fExtend = LegacyDocFib.ReadUInt16(tableStream, offset);
            ushort count = LegacyDocFib.ReadUInt16(tableStream, offset + 2);
            ushort cbExtra = LegacyDocFib.ReadUInt16(tableStream, offset + 4);
            if (fExtend != 0xFFFF || cbExtra != 0) {
                warning = "The DOC bookmark name table is not an extended string table without extra data.";
                return false;
            }

            var result = new List<string>(count);
            int cursor = offset + 6;
            int endOffset = offset + length;
            var seen = new HashSet<string>(StringComparer.Ordinal);
            for (int index = 0; index < count; index++) {
                if (cursor + 2 > endOffset) {
                    warning = "The DOC bookmark name table ended before all names were read.";
                    return false;
                }

                ushort characterCount = LegacyDocFib.ReadUInt16(tableStream, cursor);
                cursor += 2;
                if (characterCount == 0 || characterCount > 40) {
                    warning = "The DOC bookmark name table contains an invalid bookmark name length.";
                    return false;
                }

                int byteCount = checked(characterCount * 2);
                if (cursor + byteCount > endOffset) {
                    warning = "The DOC bookmark name table contains a truncated bookmark name.";
                    return false;
                }

                string name = Encoding.Unicode.GetString(tableStream, cursor, byteCount);
                cursor += byteCount;
                if (!seen.Add(name)) {
                    warning = "The DOC bookmark name table contains duplicate bookmark names.";
                    return false;
                }

                result.Add(name);
            }

            names = result;
            return true;
        }

        private static bool IsRangeWithin(byte[] bytes, int offset, int length) {
            return offset >= 0
                && length >= 0
                && offset <= bytes.Length
                && length <= bytes.Length - offset;
        }
    }
}
