using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private static void AddLegacyDocRevisionRunContent(
            WordParagraph paragraph,
            LegacyDocTextRun legacyRun,
            LegacyDocNoteProjection notes,
            LegacyDocBookmarkProjection bookmarks) {
            string text = legacyRun.Text;
            int segmentStart = 0;
            for (int index = 0; index < text.Length; index++) {
                char character = text[index];
                int? markerPosition = GetLegacyDocRunCharacterPosition(legacyRun, index);
                if (!IsLegacyDocSpecialRunCharacter(character)) {
                    if (!bookmarks.HasMarkers(markerPosition)) {
                        continue;
                    }

                    AddLegacyDocRevisionTextSegment(paragraph, legacyRun, segmentStart, index - segmentStart);
                    bookmarks.EmitAt(paragraph._paragraph, markerPosition);
                    segmentStart = index;
                    continue;
                }

                AddLegacyDocRevisionTextSegment(paragraph, legacyRun, segmentStart, index - segmentStart);
                bookmarks.EmitAt(paragraph._paragraph, markerPosition);
                if (character == '\t') {
                    AddLegacyDocRevisionElement(paragraph, legacyRun, new TabChar());
                } else if (character == LegacyDocFootnoteReader.FootnoteReferenceCharacter) {
                    AddLegacyDocNoteReference(paragraph, notes, markerPosition);
                } else if (character == LegacyDocCommentReader.CommentReferenceCharacter) {
                    AddLegacyDocCommentReference(paragraph, notes, markerPosition);
                } else {
                    AddLegacyDocRevisionElement(paragraph, legacyRun, new Break { Type = GetLegacyDocBreakType(character) });
                }

                segmentStart = index + 1;
            }

            AddLegacyDocRevisionTextSegment(paragraph, legacyRun, segmentStart, text.Length - segmentStart);
            bookmarks.EmitAt(paragraph._paragraph, GetLegacyDocRunEndCharacterPosition(legacyRun));
        }

        private static void AddLegacyDocRevisionTextSegment(WordParagraph paragraph, LegacyDocTextRun legacyRun, int startIndex, int length) {
            if (length <= 0) {
                return;
            }

            string value = legacyRun.Text.Substring(startIndex, length);
            OpenXmlLeafTextElement textNode = legacyRun.Revision.Kind == LegacyDocRevisionKind.Deleted
                ? new DeletedText(value) { Space = SpaceProcessingModeValues.Preserve }
                : new Text(value) { Space = SpaceProcessingModeValues.Preserve };
            AddLegacyDocRevisionElement(paragraph, legacyRun, textNode);
        }

        private static void AddLegacyDocRevisionElement(WordParagraph paragraph, LegacyDocTextRun legacyRun, OpenXmlElement content) {
            var run = new Run(content);
            ApplyLegacyDocRunFormatting(new WordParagraph(paragraph._document, paragraph._paragraph, run), legacyRun);
            AppendLegacyDocRevisionRun(paragraph, legacyRun, run);
        }

        private static void AppendLegacyDocRevisionRun(WordParagraph paragraph, LegacyDocTextRun legacyRun, Run run) {
            OpenXmlCompositeElement revisionElement = legacyRun.Revision.Kind == LegacyDocRevisionKind.Deleted
                ? new DeletedRun()
                : new InsertedRun();

            if (revisionElement is DeletedRun deletedRun) {
                deletedRun.Author = legacyRun.Revision.Author ?? LegacyDocRevisionAuthorReader.UnknownAuthor;
                deletedRun.Id = WordHeadersAndFooters.GenerateRevisionId();
                if (legacyRun.Revision.Date != null) {
                    deletedRun.Date = legacyRun.Revision.Date.Value;
                }
            } else if (revisionElement is InsertedRun insertedRun) {
                insertedRun.Author = legacyRun.Revision.Author ?? LegacyDocRevisionAuthorReader.UnknownAuthor;
                insertedRun.Id = WordHeadersAndFooters.GenerateRevisionId();
                if (legacyRun.Revision.Date != null) {
                    insertedRun.Date = legacyRun.Revision.Date.Value;
                }
            }

            revisionElement.Append(run);
            paragraph._paragraph.Append(revisionElement);
        }
    }
}
