namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocParagraphFormattingWriter {
        internal static IReadOnlyList<IReadOnlyList<LegacyDocWritableParagraphSegment>> CreatePapxFkpPages(
            IReadOnlyList<LegacyDocWritableParagraphSegment> segments,
            int pageSize) {
            if (segments.Count == 0) {
                return Array.Empty<IReadOnlyList<LegacyDocWritableParagraphSegment>>();
            }

            var pages = new List<IReadOnlyList<LegacyDocWritableParagraphSegment>>();
            var currentPage = new List<LegacyDocWritableParagraphSegment>();
            foreach (LegacyDocWritableParagraphSegment segment in segments) {
                currentPage.Add(segment);
                if (CanFitPapxPage(currentPage, pageSize)) {
                    continue;
                }

                currentPage.RemoveAt(currentPage.Count - 1);
                if (currentPage.Count == 0) {
                    throw new NotSupportedException("Native DOC saving encountered a paragraph-format run that cannot fit in a paragraph-format page.");
                }

                pages.Add(currentPage.ToArray());
                currentPage.Clear();
                currentPage.Add(segment);
                if (!CanFitPapxPage(currentPage, pageSize)) {
                    throw new NotSupportedException("Native DOC saving encountered a paragraph-format run that cannot fit in a paragraph-format page.");
                }
            }

            if (currentPage.Count > 0) {
                pages.Add(currentPage.ToArray());
            }

            return pages;
        }

        private static bool CanFitPapxPage(IReadOnlyList<LegacyDocWritableParagraphSegment> segments, int pageSize) {
            if (segments.Count == 0 || segments.Count > byte.MaxValue) {
                return false;
            }

            int dataBoundary = ((segments.Count + 1) * sizeof(int)) + (segments.Count * PapxFkpBxLength);
            int papxOffset = pageSize - 1;
            if (dataBoundary >= papxOffset) {
                return false;
            }

            foreach (LegacyDocWritableParagraphSegment segment in segments) {
                byte[]? papx = segment.PapxOverride;
                if (papx == null && segment.Formatting.HasFormatting) {
                    papx = CreatePapx(segment.Formatting);
                }

                if (papx == null) {
                    continue;
                }

                papxOffset -= papx.Length;
                papxOffset = papxOffset % 2 == 0 ? papxOffset : papxOffset - 1;
                if (papxOffset <= dataBoundary || papxOffset / 2 > byte.MaxValue) {
                    return false;
                }
            }

            return true;
        }
    }
}
