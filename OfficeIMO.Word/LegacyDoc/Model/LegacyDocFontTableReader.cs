using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocFontTableReader {
        private const int FfnNameOffset = 39;

        internal static IReadOnlyList<string> ReadFontFamilies(byte[] tableStream, LegacyDocFib fib, out string? warning) {
            warning = null;
            if (fib.LcbSttbfFfn == 0) {
                return Array.Empty<string>();
            }

            if (fib.LcbSttbfFfn < 4
                || fib.FcSttbfFfn < 0
                || fib.FcSttbfFfn + fib.LcbSttbfFfn > tableStream.Length) {
                warning = "The FIB points outside the selected table stream for the font table.";
                return Array.Empty<string>();
            }

            int offset = fib.FcSttbfFfn;
            int end = offset + fib.LcbSttbfFfn;
            ushort first = LegacyDocFib.ReadUInt16(tableStream, offset);
            if (first == 0xFFFF) {
                warning = "Extended DOC font tables are detected but are not imported by the dependency-free reader yet.";
                return Array.Empty<string>();
            }

            int count = first;
            offset += 2;
            int cbExtra = LegacyDocFib.ReadUInt16(tableStream, offset);
            offset += 2;

            var fonts = new List<string>(count);
            for (int index = 0; index < count; index++) {
                if (offset >= end) {
                    warning = "The DOC font table ended before all declared font records were read.";
                    break;
                }

                int cbFfn = tableStream[offset];
                offset++;
                if (cbFfn <= 0 || offset + cbFfn > end) {
                    warning = "A DOC font table record points outside the font table.";
                    break;
                }

                fonts.Add(ReadFfnName(tableStream, offset, cbFfn) ?? string.Empty);
                offset += cbFfn;
                if (cbExtra > 0) {
                    if (offset + cbExtra > end) {
                        warning = "A DOC font table extra-data record points outside the font table.";
                        break;
                    }

                    offset += cbExtra;
                }
            }

            return fonts;
        }

        private static string? ReadFfnName(byte[] bytes, int offset, int count) {
            int nameOffset = offset + FfnNameOffset;
            int end = offset + count;
            if (nameOffset + 2 > end) {
                return null;
            }

            int nameByteCount = 0;
            for (int position = nameOffset; position + 1 < end; position += 2) {
                if (bytes[position] == 0 && bytes[position + 1] == 0) {
                    break;
                }

                nameByteCount += 2;
            }

            if (nameByteCount == 0) {
                return null;
            }

            return Encoding.Unicode.GetString(bytes, nameOffset, nameByteCount);
        }
    }
}
