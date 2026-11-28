using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    // SDK-side paragraph property reads for code that walks SDK paragraphs directly, such as field and table of contents
    // updates.
    public partial class WordParagraph {
        internal static bool HasPageBreakBefore(Paragraph paragraph) {
            PageBreakBefore? value = paragraph.GetFirstChild<ParagraphProperties>()?.PageBreakBefore;
            return value?.Val?.Value ?? value is not null;
        }
    }
}
