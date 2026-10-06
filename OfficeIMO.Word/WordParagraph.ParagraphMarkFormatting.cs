using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Copies this run's direct formatting to the paragraph mark. Importers use
        /// this after applying paragraph text defaults so empty paragraphs retain
        /// their source font metrics instead of the destination document defaults.
        /// </summary>
        internal void CopyRunFormattingToParagraphMark() {
            RunProperties? runProperties = _runProperties;
            if (runProperties == null) return;

            var markProperties = new ParagraphMarkRunProperties();
            foreach (OpenXmlElement property in runProperties.ChildElements) {
                markProperties.Append(property.CloneNode(true));
            }
            ParagraphProperties properties = _paragraph.ParagraphProperties ??= new ParagraphProperties();
            properties.ParagraphMarkRunProperties = markProperties;
        }
    }
}
