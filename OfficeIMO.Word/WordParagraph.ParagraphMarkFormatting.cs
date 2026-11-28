using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Copies this run's direct formatting to the paragraph mark. Importers use
        /// this after applying paragraph text defaults so empty paragraphs retain
        /// their source font metrics instead of the destination document defaults.
        /// </summary>
        internal void CopyRunFormattingToParagraphMark() {
            OfficeXmlNode runProperties = OwnRunPropertiesNode;
            if (runProperties.IsNone) return;

            OfficeXmlNode mark = WordXmlEdits.ReplaceChild(EditParagraphProperties(), "rPr", WPPrView.Order);
            foreach (OfficeXmlNode property in runProperties.Children()) {
                mark.Append(property.Clone());
            }
        }
    }
}
