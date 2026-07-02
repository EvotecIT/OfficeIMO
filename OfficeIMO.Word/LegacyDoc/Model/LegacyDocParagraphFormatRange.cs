namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocParagraphFormatRange {
        internal LegacyDocParagraphFormatRange(int fileOffsetStart, int fileOffsetEnd, LegacyDocParagraphFormat format) {
            FileOffsetStart = fileOffsetStart;
            FileOffsetEnd = fileOffsetEnd;
            Format = format;
        }

        internal int FileOffsetStart { get; }

        internal int FileOffsetEnd { get; }

        internal LegacyDocParagraphFormat Format { get; }

        internal bool Contains(int fileOffset) {
            return fileOffset >= FileOffsetStart && fileOffset < FileOffsetEnd;
        }
    }
}
