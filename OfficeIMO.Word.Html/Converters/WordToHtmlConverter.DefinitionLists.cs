using OfficeIMO.Word;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static bool IsDefinitionListParagraph(WordParagraph paragraph) {
            return IsDefinitionTerm(paragraph) || IsDefinitionDescription(paragraph);
        }

        private static string GetDefinitionListTagName(WordParagraph paragraph) {
            return IsDefinitionDescription(paragraph) ? "dd" : "dt";
        }

        private static bool IsEmptyDefinitionListParagraph(WordParagraph paragraph) {
            return !paragraph.GetFormattedRuns()
                .Any(run => !string.IsNullOrEmpty(run.Text) || run.Image != null);
        }

        // A class-styled term or description imported in this session keeps its role through an annotation.
        private static bool IsDefinitionTerm(WordParagraph paragraph) {
            return string.Equals(paragraph.StyleId, HtmlSemanticStyleIds.DefinitionTerm, StringComparison.Ordinal) ||
                HtmlSemanticMetadata.GetDefinitionListItemRole(paragraph) == false;
        }

        private static bool IsDefinitionDescription(WordParagraph paragraph) {
            return string.Equals(paragraph.StyleId, HtmlSemanticStyleIds.DefinitionDescription, StringComparison.Ordinal) ||
                HtmlSemanticMetadata.GetDefinitionListItemRole(paragraph) == true;
        }
    }
}
