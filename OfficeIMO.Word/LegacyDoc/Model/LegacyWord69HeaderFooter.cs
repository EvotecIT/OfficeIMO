using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>Reads the variable header/footer slots used by Word 6/95.</summary>
    internal static class LegacyWord69HeaderFooter {
        internal static IReadOnlyList<LegacyDocHeaderFooterStory> Read(
            LegacyWord69Bytes positions, LegacyDocTextCharacter[] characters, int headerStart, int headerLength,
            IReadOnlyList<byte> sectionMasks, byte separatorMask,
            IReadOnlyList<LegacyDocCharacterFormatRange> characterRanges,
            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphRanges, out bool unsupported,
            LegacyDocFieldMetadataReader? fieldMetadata = null) {
            unsupported = false;
            var stories = new List<LegacyDocHeaderFooterStory>();
            if (headerLength == 0) return stories;
            if (positions.IsEmpty) {
                unsupported = true;
                return stories;
            }
            if (positions.Length < 8 || positions.Length % 4 != 0) {
                throw new InvalidDataException("The Word 6/95 header/footer directory has an invalid length.");
            }

            // Only present note separator stories occupy positions in the older PLC.
            int slot = 0;
            for (int bit = 1; bit <= 0x20; bit <<= 1) {
                if ((separatorMask & bit) != 0) slot++;
            }
            int storyCount = positions.Length / 4 - 2;
            int previousEnd = 0;
            for (int index = 0; index <= storyCount; index++) {
                int position = LegacyWord69Reader.I32(positions, index * 4);
                if (position < previousEnd || position > headerLength) {
                    throw new InvalidDataException("The Word 6/95 header/footer directory contains an invalid text range.");
                }
                previousEnd = position;
            }
            for (int section = 0; section < sectionMasks.Count; section++) {
                for (int index = 0; index < 6; index++) {
                    if ((sectionMasks[section] & (1 << index)) == 0) continue;
                    if (slot >= storyCount) {
                        throw new InvalidDataException("The Word 6/95 header/footer directory omits a declared story.");
                    }
                    int start = LegacyWord69Reader.I32(positions, slot * 4);
                    int end = LegacyWord69Reader.I32(positions, (slot + 1) * 4);
                    if (start < 0 || end < start || end > headerLength) {
                        throw new InvalidDataException("A Word 6/95 header/footer range exceeds its text story.");
                    }
                    slot++;
                    // The last paragraph mark is the story guard, outside the authored content.
                    if (end > start && characters[headerStart + end - 1].Character == '\r') end--;
                    var text = new ArraySegment<LegacyDocTextCharacter>(characters, headerStart + start, end - start);
                    var bookmarks = new LegacyDocBookmarkProjectionTracker(Array.Empty<LegacyDocBookmark>());
                    var blocks = LegacyDocDocument.BuildStoryBlocks(text, characterRanges, paragraphRanges,
                        bookmarks, new Dictionary<int, LegacyDocPicture>(), orderedFormattingRanges: true, fieldMetadata: fieldMetadata);
                    if (blocks.Count == 0) continue;
                    HeaderFooterValues type = index >= 4 ? HeaderFooterValues.First
                        : index is 0 or 2 ? HeaderFooterValues.Even : HeaderFooterValues.Default;
                    stories.Add(new LegacyDocHeaderFooterStory(section, index is 0 or 1 or 4, type, blocks));
                }
            }
            unsupported = slot != storyCount || separatorMask != 0 || headerLength - previousEnd > 1;
            return stories;
        }
    }
}
