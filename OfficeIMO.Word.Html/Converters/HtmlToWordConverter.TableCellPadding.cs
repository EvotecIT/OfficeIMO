namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private void ApplyTableCellPadding(WordTableCell cell, AngleSharp.Dom.IElement element) {
            string? style = element.GetAttribute("style");
            if (string.IsNullOrWhiteSpace(style) || style!.IndexOf("padding", StringComparison.OrdinalIgnoreCase) < 0) return;
            CssStyleMapper.CssProperties parsed = ParseElementBoxStyles(element);
            if (parsed.PaddingTop.HasValue) cell.MarginTopWidth = NormalizeTableCellPadding(parsed.PaddingTop.Value, "top");
            if (parsed.PaddingRight.HasValue) cell.MarginRightWidth = NormalizeTableCellPadding(parsed.PaddingRight.Value, "right");
            if (parsed.PaddingBottom.HasValue) cell.MarginBottomWidth = NormalizeTableCellPadding(parsed.PaddingBottom.Value, "bottom");
            if (parsed.PaddingLeft.HasValue) cell.MarginLeftWidth = NormalizeTableCellPadding(parsed.PaddingLeft.Value, "left");
        }

        private short NormalizeTableCellPadding(int twips, string side) {
            if (twips > short.MaxValue) {
                AddUnsupportedCssDiagnostic("UnsupportedCssValue", "Table cell padding exceeded the native margin limit and was bounded.",
                    "td:padding-" + side, "maximum-twips=" + short.MaxValue);
            }
            return (short)Math.Max(0, Math.Min(short.MaxValue, twips));
        }
    }
}
