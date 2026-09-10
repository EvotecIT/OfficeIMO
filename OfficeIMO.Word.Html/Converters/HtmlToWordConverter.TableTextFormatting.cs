using AngleSharp.Dom;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        /// <summary>Preserves inherited table text styles without painting cell backgrounds onto individual runs.</summary>
        private TextFormatting GetTableTextFormatting(IElement element, TextFormatting inherited) {
            var formatting = inherited;
            ApplySpanStyles(element, ref formatting);
            PreserveBlockBackgroundAsTextBackdrop(ref formatting, inherited);
            return formatting;
        }
    }
}
