using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>Recognizes a section terminator with no body content, while retaining its editable properties and anchors.</summary>
        internal static bool IsSectionMarkOnly(Paragraph paragraph) =>
            paragraph.ParagraphProperties?.SectionProperties != null &&
            paragraph.ChildElements.All(element => element is ParagraphProperties or BookmarkStart or BookmarkEnd or ProofError);
    }
}
