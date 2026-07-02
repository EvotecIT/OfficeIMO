namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocCharacterFormatRange {
        internal LegacyDocCharacterFormatRange(int fileOffsetStart, int fileOffsetEnd, LegacyDocCharacterFormat format) {
            FileOffsetStart = fileOffsetStart;
            FileOffsetEnd = fileOffsetEnd;
            Format = format;
        }

        internal int FileOffsetStart { get; }

        internal int FileOffsetEnd { get; }

        internal LegacyDocCharacterFormat Format { get; }

        internal bool Contains(int fileOffset) {
            return fileOffset >= FileOffsetStart && fileOffset < FileOffsetEnd;
        }
    }
}
