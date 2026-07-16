namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void WriteChpxBtePlc(
            byte[] table,
            LegacyDocWritableBody body,
            int chpxFkpOffset,
            int bytesPerCharacter) {
            IReadOnlyList<IReadOnlyList<LegacyDocWritableSegment>> pages = body.ChpxPages;
            int pageNumbersOffset = ClxLength + ((pages.Count + 1) * sizeof(int));
            for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
                WriteInt32(
                    table,
                    ClxLength + (pageIndex * sizeof(int)),
                    TextOffset + (pages[pageIndex][0].StartCharacter * bytesPerCharacter));
                WriteInt32(
                    table,
                    pageNumbersOffset + (pageIndex * sizeof(int)),
                    (chpxFkpOffset / OleSectorSize) + pageIndex);
            }

            IReadOnlyList<LegacyDocWritableSegment> lastPage = pages[pages.Count - 1];
            WriteInt32(
                table,
                ClxLength + (pages.Count * sizeof(int)),
                TextOffset + (lastPage[lastPage.Count - 1].EndCharacter * bytesPerCharacter));
        }

        private static void WritePapxBtePlc(
            byte[] table,
            LegacyDocWritableBody body,
            int papxFkpOffset,
            int bytesPerCharacter) {
            IReadOnlyList<IReadOnlyList<LegacyDocWritableParagraphSegment>> pages = body.PapxPages;
            int pageNumbersOffset = body.PapxPlcOffsetInTableStream + ((pages.Count + 1) * sizeof(int));
            for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
                WriteInt32(
                    table,
                    body.PapxPlcOffsetInTableStream + (pageIndex * sizeof(int)),
                    TextOffset + (pages[pageIndex][0].StartCharacter * bytesPerCharacter));
                WriteInt32(
                    table,
                    pageNumbersOffset + (pageIndex * sizeof(int)),
                    (papxFkpOffset / OleSectorSize) + pageIndex);
            }

            IReadOnlyList<LegacyDocWritableParagraphSegment> lastPage = pages[pages.Count - 1];
            WriteInt32(
                table,
                body.PapxPlcOffsetInTableStream + (pages.Count * sizeof(int)),
                TextOffset + (lastPage[lastPage.Count - 1].EndCharacter * bytesPerCharacter));
        }
    }
}
