namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>Reads Word 6/95 bin tables and the older 512-byte formatting pages.</summary>
    internal static class LegacyWord69Formatting {
        private const int PageSize = 512;

        internal static IReadOnlyList<LegacyDocCharacterFormatRange> ReadCharacterRanges(
            byte[] bytes, int usedEnd, ReadOnlySpan<byte> bin, IReadOnlyList<string> fonts,
            LegacyWord69Properties properties, int maximumWork, out IReadOnlyList<LegacyWord69Encoding.TextContextRange> contextRanges) {
            var ranges = new List<LegacyDocCharacterFormatRange>();
            var textContexts = new List<LegacyWord69Encoding.TextContextRange>();
            contextRanges = textContexts;
            foreach (var run in ReadRuns(bytes, usedEnd, bin, paragraph: false, maximumWork)) {
                if (run.Length == 0) continue;
                var format = properties.ReadCharacter(bytes.AsSpan(run.Offset, run.Length), fonts, out int? fontIndex, out ushort? languageId);
                if (fontIndex.HasValue || languageId.HasValue) {
                    textContexts.Add(new LegacyWord69Encoding.TextContextRange(run.Start, run.End, fontIndex, languageId));
                }
                if (format.HasFormatting) ranges.Add(new LegacyDocCharacterFormatRange(run.Start, run.End, format));
            }

            return ranges;
        }

        internal static IReadOnlyList<LegacyDocParagraphFormatRange> ReadParagraphRanges(
            byte[] bytes, int usedEnd, ReadOnlySpan<byte> bin, LegacyWord69Properties properties, int maximumWork) {
            var ranges = new List<LegacyDocParagraphFormatRange>();
            foreach (var run in ReadRuns(bytes, usedEnd, bin, paragraph: true, maximumWork)) {
                if (run.Length == 0) continue;
                var data = bytes.AsSpan(run.Offset, run.Length);
                var format = properties.ReadParagraph(data.Slice(2), LegacyWord69Reader.U16(data, 0));
                ranges.Add(new LegacyDocParagraphFormatRange(run.Start, run.End, format));
            }

            return ranges;
        }

        private static List<FormattingRun> ReadRuns(
            byte[] bytes, int usedEnd, ReadOnlySpan<byte> bin, bool paragraph, int maximumWork) {
            var runs = new List<FormattingRun>();
            if (bin.IsEmpty) return runs;
            if (bin.Length < 10 || (bin.Length - 4) % 6 != 0) {
                throw new InvalidDataException("A Word 6/95 formatting bin table has an invalid record length.");
            }

            int pageCount = (bin.Length - 4) / 6;
            int pageIndex = (pageCount + 1) * 4;
            int stride = paragraph ? 7 : 1;
            int lastEnd = 0;
            var seen = new HashSet<int>();
            for (int i = 0; i < pageCount; i++) {
                int pageOffset = LegacyWord69Reader.U16(bin, pageIndex + 2 * i) * PageSize;
                if (pageOffset < PageSize || pageOffset > usedEnd - PageSize || !seen.Add(pageOffset)) {
                    throw new InvalidDataException("A Word 6/95 formatting bin contains a repeated or invalid page reference.");
                }

                ReadOnlySpan<byte> page = bytes.AsSpan(pageOffset, PageSize);
                int count = page[PageSize - 1];
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
