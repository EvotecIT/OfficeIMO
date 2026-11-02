namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private void ApplyTableCellPadding(WordTableCell cell, AngleSharp.Dom.IElement element) {
            string? style = element.GetAttribute("style");
            if (string.IsNullOrWhiteSpace(style) || style!.IndexOf("padding", StringComparison.OrdinalIgnoreCase) < 0) return;
            CssStyleMapper.CssProperties parsed = ParseElementBoxStyles(element);
            if (parsed.PaddingTop.HasValue) cell.MarginTop = global::OfficeIMO.OfficeLength.Points((NormalizeTableCellPadding(parsed.PaddingTop.Value, "top")) / 20D);
            if (parsed.PaddingRight.HasValue) cell.MarginRight = global::OfficeIMO.OfficeLength.Points((NormalizeTableCellPadding(parsed.PaddingRight.Value, "right")) / 20D);
            if (parsed.PaddingBottom.HasValue) cell.MarginBottom = global::OfficeIMO.OfficeLength.Points((NormalizeTableCellPadding(parsed.PaddingBottom.Value, "bottom")) / 20D);
            if (parsed.PaddingLeft.HasValue) cell.MarginLeft = global::OfficeIMO.OfficeLength.Points((NormalizeTableCellPadding(parsed.PaddingLeft.Value, "left")) / 20D);
        }

        private short NormalizeTableCellPadding(int twips, string side) {
            if (twips > short.MaxValue) {
                AddUnsupportedCssDiagnostic("unsupported-css-value", "Table cell padding exceeded the native margin limit and was bounded.",
                    "td:padding-" + side, "maximum-twips=" + short.MaxValue);
            }
            return (short)Math.Max(0, Math.Min(short.MaxValue, twips));
        }
    }
}
