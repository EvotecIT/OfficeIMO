using System.Globalization;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {

        /// <summary>
        /// Header-row appearance taken from the table style's <c>firstRow</c> conditional formatting.
        /// </summary>
        private sealed class TableFirstRowAppearance {
            internal bool Bold { get; set; }
            internal string? Color { get; set; }
            internal IEnumerable<string> ToCss() {
                if (Bold) yield return "font-weight:bold";
                if (Color != null) yield return "color:#" + Color;
            }
        }

        /// <summary>
        /// Returns the header appearance when the table's look enables first-row formatting and its
        /// style (or a style it is based on) defines it; otherwise null.
        /// </summary>
        private static TableFirstRowAppearance? ResolveTableFirstRowAppearance(WordTable table) {
            if (table.Rows.Count < 2 || !IsTableFirstRowLookEnabled(table.Node())) {
                return null;
            }

            string? styleId = table.Node().WChild("tblPr").WChild("tblStyle").WVal();
            if (string.IsNullOrWhiteSpace(styleId)) {
                return null;
            }

            if (table.Document == null) {
                return null;
            }

            var conditionals = WordStyleCatalog.For(table.Document).StyleChainNodes(styleId, "table")
                .SelectMany(style => style.WElements("tblStylePr"))
                .Where(properties => properties.WAttr("type") == "firstRow")
                .ToList();
            if (conditionals.Count == 0) {
                return null;
            }

            var appearance = new TableFirstRowAppearance();
            bool? bold = null;
            bool colorSpecified = false;
            foreach (OfficeXmlNode conditional in conditionals) {
                OfficeXmlNode run = conditional.WChild("rPr");
                if (bold == null && run.WChild("b") is { Exists: true } boldElement) {
                    bold = boldElement.IsOn();
                }
                if (!colorSpecified && run.WChild("color") is { Exists: true } color) {
                    appearance.Color = NormalizeStyleColor(WordConverterNodes.ResolveColorHex(color, table.Document));
                    colorSpecified = true;
                }
            }
            appearance.Bold = bold == true;
            return appearance;
        }

        private static bool IsTableFirstRowLookEnabled(OfficeXmlNode table) => WordConverterNodes.HasFirstRowFormatting(table);

        private static string? NormalizeStyleColor(string? value) {
            if (string.IsNullOrWhiteSpace(value) || string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase)) {
                return null;
            }
            string trimmed = value!.Trim().TrimStart('#');
            return trimmed.Length == 6 && trimmed.All(Uri.IsHexDigit) ? trimmed.ToUpperInvariant() : null;
        }

    }
}
