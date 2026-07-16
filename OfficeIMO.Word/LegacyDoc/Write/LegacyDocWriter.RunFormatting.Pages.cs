namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static IReadOnlyList<IReadOnlyList<LegacyDocWritableSegment>> CreateChpxFkpPages(
            IReadOnlyList<LegacyDocWritableSegment> segments,
            IReadOnlyDictionary<string, int> fontFamilyIndexes,
            IReadOnlyDictionary<string, int> revisionAuthorIndexes) {
            if (segments.Count == 0) {
                return Array.Empty<IReadOnlyList<LegacyDocWritableSegment>>();
            }

            var pages = new List<IReadOnlyList<LegacyDocWritableSegment>>();
            var currentPage = new List<LegacyDocWritableSegment>();
            foreach (LegacyDocWritableSegment segment in segments) {
                currentPage.Add(segment);
                if (CanFitChpxPage(currentPage, fontFamilyIndexes, revisionAuthorIndexes)) {
                    continue;
                }

                currentPage.RemoveAt(currentPage.Count - 1);
                if (currentPage.Count == 0) {
                    throw new NotSupportedException("Native DOC saving encountered a character-format run that cannot fit in a character-format page.");
                }

                pages.Add(currentPage.ToArray());
                currentPage.Clear();
                currentPage.Add(segment);
                if (!CanFitChpxPage(currentPage, fontFamilyIndexes, revisionAuthorIndexes)) {
                    throw new NotSupportedException("Native DOC saving encountered a character-format run that cannot fit in a character-format page.");
                }
            }

            if (currentPage.Count > 0) {
                pages.Add(currentPage.ToArray());
            }

            return pages;
        }

        private static bool CanFitChpxPage(
            IReadOnlyList<LegacyDocWritableSegment> segments,
            IReadOnlyDictionary<string, int> fontFamilyIndexes,
            IReadOnlyDictionary<string, int> revisionAuthorIndexes) {
            if (segments.Count == 0 || segments.Count > byte.MaxValue) {
                return false;
            }

            int chpxOffset = AlignToEven(((segments.Count + 1) * sizeof(int)) + segments.Count);
            if (chpxOffset >= OleSectorSize - 1 || chpxOffset / 2 > byte.MaxValue) {
                return false;
            }

            foreach (LegacyDocWritableSegment segment in segments) {
                if (!segment.HasFormatting) {
                    continue;
                }

                byte[] chpx = CreateChpx(segment.Formatting, fontFamilyIndexes, revisionAuthorIndexes, segment.PictureDataOffset);
                chpxOffset = AlignToEven(chpxOffset);
                if (chpxOffset + chpx.Length >= OleSectorSize - 1 || chpxOffset / 2 > byte.MaxValue) {
                    return false;
                }

                chpxOffset += chpx.Length;
            }

            return true;
        }
    }
}
