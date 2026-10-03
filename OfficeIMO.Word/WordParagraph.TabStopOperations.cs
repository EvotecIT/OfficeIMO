using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>Removes this paragraph's explicit tab stops, leaving inherited defaults unchanged.</summary>
        public WordParagraph ClearTabStops() {
            _paragraph.ParagraphProperties?.RemoveAllChildren<Tabs>();
            return this;
        }
    }
}
