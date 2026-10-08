using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private static void AddLegacyDocPageNumber(WordParagraph paragraph, LegacyDocTextRun legacyRun, LegacyDocBookmarkProjection bookmarks) {
            bookmarks.EmitAt(paragraph._paragraph, GetLegacyDocRunCharacterPosition(legacyRun, 0));
            var run = new Run(new PageNumber());
            paragraph._paragraph.Append(run);
            ApplyLegacyDocRunFormatting(new WordParagraph(paragraph._document, paragraph._paragraph, run), legacyRun);
            bookmarks.EmitAt(paragraph._paragraph, GetLegacyDocRunEndCharacterPosition(legacyRun));
        }

        private static void AddLegacyDocNumberOfPages(WordParagraph paragraph, LegacyDocTextRun legacyRun, LegacyDocBookmarkProjection bookmarks) {
            AddLegacyDocPageCountField(paragraph, legacyRun, bookmarks, " NUMPAGES  ");
        }

        private static void AddLegacyDocSectionPages(WordParagraph paragraph, LegacyDocTextRun legacyRun, LegacyDocBookmarkProjection bookmarks) {
            string instruction = string.IsNullOrWhiteSpace(legacyRun.FieldInstruction) ? " SECTIONPAGES  " : legacyRun.FieldInstruction!;
            AddLegacyDocPageCountField(paragraph, legacyRun, bookmarks, instruction);
        }

        // Import retains the producer's cached count; the field remains available for later layout/evaluation.
        private static void AddLegacyDocPageCountField(WordParagraph paragraph, LegacyDocTextRun legacyRun, LegacyDocBookmarkProjection bookmarks, string instruction) {
            bookmarks.EmitAt(paragraph._paragraph, GetLegacyDocRunCharacterPosition(legacyRun, 0));
            var simpleField = new SimpleField { Instruction = instruction };
            AppendLegacyDocFieldResultContent(simpleField, paragraph, legacyRun, string.IsNullOrEmpty(legacyRun.Text) ? "1" : legacyRun.Text);
            paragraph._paragraph.Append(simpleField);
            bookmarks.EmitAt(paragraph._paragraph, GetLegacyDocRunEndCharacterPosition(legacyRun));
        }
    }
}
