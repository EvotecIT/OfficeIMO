using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocRevisionAuthorReader {
        internal const string UnknownAuthor = "Unknown";

        internal static IReadOnlyList<string> Read(byte[] tableStream, LegacyDocFib fib, out string? warning) {
            warning = null;
            if (fib.LcbSttbfRMark == 0) {
                return Array.Empty<string>();
            }

            if (!IsRangeWithin(tableStream, fib.FcSttbfRMark, fib.LcbSttbfRMark)
                || fib.LcbSttbfRMark < 6) {
                warning = "The DOC revision-author string table extends beyond the selected table stream.";
                return Array.Empty<string>();
            }

            int cursor = fib.FcSttbfRMark;
            int end = cursor + fib.LcbSttbfRMark;
            ushort fExtend = LegacyDocFib.ReadUInt16(tableStream, cursor);
            ushort count = LegacyDocFib.ReadUInt16(tableStream, cursor + 2);
            ushort cbExtra = LegacyDocFib.ReadUInt16(tableStream, cursor + 4);
            if (fExtend != 0xFFFF || cbExtra != 0) {
                warning = "The DOC revision-author table is not an extended Unicode string table without extra data.";
                return Array.Empty<string>();
            }

            cursor += 6;
            var authors = new List<string>(count);
            for (int index = 0; index < count; index++) {
                if (cursor + 2 > end) {
                    warning = "The DOC revision-author table ended before all declared names were read.";
                    return Array.Empty<string>();
                }

                int characterCount = LegacyDocFib.ReadUInt16(tableStream, cursor);
                cursor += 2;
                int byteCount = checked(characterCount * 2);
                if (cursor + byteCount > end) {
                    warning = "The DOC revision-author table contains a truncated name.";
                    return Array.Empty<string>();
                }

                authors.Add(Encoding.Unicode.GetString(tableStream, cursor, byteCount));
                cursor += byteCount;
            }

            if (authors.Count == 0
                || !string.Equals(authors[0], UnknownAuthor, StringComparison.Ordinal)) {
                warning = "The DOC revision-author table does not begin with the required Unknown author entry.";
                return Array.Empty<string>();
            }

            return authors;
        }

        private static bool IsRangeWithin(byte[] bytes, int offset, int length) {
            return offset >= 0
                && length >= 0
                && offset <= bytes.Length
                && length <= bytes.Length - offset;
        }
    }
}
