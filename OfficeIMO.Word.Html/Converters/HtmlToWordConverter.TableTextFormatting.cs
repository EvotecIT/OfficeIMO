using AngleSharp.Dom;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        /// <summary>Preserves inherited table text styles without painting cell backgrounds onto individual runs.</summary>
        private TextFormatting GetTableTextFormatting(IElement element, TextFormatting inherited) {
            var formatting = inherited;
            if (element.TagName.Equals("th", StringComparison.OrdinalIgnoreCase)) formatting.Bold = true;
            ApplySpanStyles(element, ref formatting);
            PreserveBlockBackgroundAsTextBackdrop(ref formatting, inherited);
            return formatting;
        }
    }
}
