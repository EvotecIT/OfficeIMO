namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>Reads Word 6/95 bin tables and the older 512-byte formatting pages.</summary>
    internal static class LegacyWord69Formatting {
        private const int PageSize = 512;

        internal static IReadOnlyList<LegacyDocCharacterFormatRange> ReadCharacterRanges(
            LegacyWord69Bytes bytes, int usedEnd, LegacyWord69Bytes bin, IReadOnlyList<string> fonts,
            LegacyWord69Properties properties, int maximumWork, out IReadOnlyList<LegacyWord69Encoding.TextContextRange> contextRanges) {
            var ranges = new List<LegacyDocCharacterFormatRange>();
            var textContexts = new List<LegacyWord69Encoding.TextContextRange>();
            contextRanges = textContexts;
            foreach (var run in ReadRuns(bytes, usedEnd, bin, paragraph: false, maximumWork)) {
                if (run.Length == 0) continue;
                var format = properties.ReadCharacter(bytes.Slice(run.Offset, run.Length), fonts, out int? fontIndex, out ushort? languageId);
                if (fontIndex.HasValue || languageId.HasValue) {
                    textContexts.Add(new LegacyWord69Encoding.TextContextRange(run.Start, run.End, fontIndex, languageId));
                }
                if (format.HasFormatting) ranges.Add(new LegacyDocCharacterFormatRange(run.Start, run.End, format));
            }

            return ranges;
        }

        internal static IReadOnlyList<LegacyDocParagraphFormatRange> ReadParagraphRanges(
            LegacyWord69Bytes bytes, int usedEnd, LegacyWord69Bytes bin, LegacyWord69Properties properties, int maximumWork) {
            var ranges = new List<LegacyDocParagraphFormatRange>();
            foreach (var run in ReadRuns(bytes, usedEnd, bin, paragraph: true, maximumWork)) {
                if (run.Length == 0) continue;
                var data = bytes.Slice(run.Offset, run.Length);
                var format = properties.ReadParagraph(data.Slice(2), LegacyWord69Reader.U16(data, 0));
                ranges.Add(new LegacyDocParagraphFormatRange(run.Start, run.End, format));
            }

            return ranges;
        }

        private static List<FormattingRun> ReadRuns(
            LegacyWord69Bytes bytes, int usedEnd, LegacyWord69Bytes bin, bool paragraph, int maximumWork) {
            var runs = new List<FormattingRun>();
            if (!bin.IsEmpty && (bin.Length < 4 || (bin.Length - 4) % 6 != 0)) {
                throw new InvalidDataException("A Word 6/95 formatting bin table has an invalid record length.");
            }

            int pageCount = bin.IsEmpty ? 0 : (bin.Length - 4) / 6;
            int pageIndex = (pageCount + 1) * 4;
            int declaredPageCount = LegacyWord69Reader.U16(bytes, paragraph ? 0x190 : 0x18E);
            bool recoverFromFib = declaredPageCount > pageCount;
            int firstPage = 0;
            if (recoverFromFib) {
                // Word 6/95 can omit bin entries while retaining contiguous formatting pages.
                firstPage = LegacyWord69Reader.U16(bytes, paragraph ? 0x18C : 0x18A);
                int lastPage = firstPage + declaredPageCount - 1;
                if (firstPage == 0 || lastPage > ushort.MaxValue || lastPage >= usedEnd / PageSize) {
                    throw new InvalidDataException("The Word 6/95 FIB declares formatting pages outside the used stream.");
                }

                pageCount = declaredPageCount;
            }
            if (pageCount == 0) return runs;
            int stride = paragraph ? 7 : 1;
            int lastEnd = 0;
            var seen = new HashSet<int>();
            for (int i = 0; i < pageCount; i++) {
                int pageNumber = recoverFromFib ? firstPage + i : LegacyWord69Reader.U16(bin, pageIndex + 2 * i);
                int pageOffset = pageNumber * PageSize;
                if (pageOffset < PageSize || pageOffset > usedEnd - PageSize || !seen.Add(pageOffset)) {
                    throw new InvalidDataException("A Word 6/95 formatting bin contains a repeated or invalid page reference.");
                }

                LegacyWord69Bytes page = bytes.Slice(pageOffset, PageSize);
                int count = page[PageSize - 1];
                if (recoverFromFib && count == 0) {
                    throw new InvalidDataException("The Word 6/95 FIB declares an empty formatting page.");
                }

                int positionsEnd = (count + 1) * 4;
                int indexEnd = positionsEnd + count * stride;
                if (indexEnd > PageSize - 1 || count > maximumWork / PageSize) {
                    throw new InvalidDataException("A Word 6/95 formatting page exceeds its layout or work budget.");
                }

                maximumWork -= count * PageSize;
                for (int j = 0; j < count; j++) {
                    int start = LegacyWord69Reader.I32(page, 4 * j);
                    int end = LegacyWord69Reader.I32(page, 4 * (j + 1));
                    if (start < lastEnd || end <= start || end > usedEnd) {
                        throw new InvalidDataException("A Word 6/95 formatting page has an invalid text range.");
                    }

                    lastEnd = end;
                    int data = page[positionsEnd + stride * j] * 2;
                    if (data == 0) continue;
                    if (data < indexEnd || data >= PageSize - 1) {
                        throw new InvalidDataException("A Word 6/95 formatting record overlaps the formatting page index.");
                    }

                    int length = paragraph ? page[data] * 2 : page[data];
                    if (length < (paragraph ? 2 : 0) || length > PageSize - 2 - data) {
                        throw new InvalidDataException("A Word 6/95 formatting record extends beyond its page.");
                    }

                    runs.Add(new FormattingRun(start, end, pageOffset + data + 1, length));
                }
            }

            return runs;
        }

        private readonly struct FormattingRun {
            internal FormattingRun(int start, int end, int offset, int length) {
                Start = start;
                End = end;
                Offset = offset;
                Length = length;
            }

            internal int Start { get; }
            internal int End { get; }
            internal int Offset { get; }
            internal int Length { get; }
        }
    }
}
