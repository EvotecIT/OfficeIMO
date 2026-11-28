namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>Removes this paragraph's explicit tab stops, leaving inherited defaults unchanged.</summary>
        public WordParagraph ClearTabStops() {
            if (!ParagraphPropertiesNode.Child(W, "tabs").IsNone) EditParagraphProperties().RemoveChildren(W, "tabs");
            return this;
        }
    }
}
