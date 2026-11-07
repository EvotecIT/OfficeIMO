using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private const int TableLookFirstRowMask = 0x0020;

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
            if (table.Rows.Count < 2 || !IsTableFirstRowLookEnabled(table._table)) {
                return null;
            }

            string? styleId = table._table.GetFirstChild<TableProperties>()?.TableStyle?.Val?.Value;
            if (string.IsNullOrWhiteSpace(styleId)) {
                return null;
            }

            Styles? styles = table.Document?._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            var conditionals = WordRunFormatting.EnumerateStyleChain(styles, styleId, StyleValues.Table)
                .SelectMany(style => style.Elements<TableStyleProperties>())
                .Where(properties => properties.Type?.Value == TableStyleOverrideValues.FirstRow)
                .ToList();
            if (conditionals.Count == 0) {
                return null;
            }

            var appearance = new TableFirstRowAppearance();
            bool? bold = null;
            bool colorSpecified = false;
            foreach (TableStyleProperties conditional in conditionals) {
                RunPropertiesBaseStyle? run = conditional.GetFirstChild<RunPropertiesBaseStyle>();
                if (bold == null && run?.GetFirstChild<Bold>() is Bold boldElement) {
                    bold = boldElement.Val == null || boldElement.Val.Value;
                }
                if (!colorSpecified && run?.GetFirstChild<Color>() is { } color) {
                    appearance.Color = NormalizeStyleColor(WordRunFormatting.ResolveColorHex(color, table.Document));
                    colorSpecified = true;
                }
            }
            appearance.Bold = bold == true;
            return appearance;
        }

        private static bool IsTableFirstRowLookEnabled(Table table) {
            TableLook? look = table.GetFirstChild<TableProperties>()?.TableLook;
            if (look == null) {
                return false;
            }
            if (look.FirstRow != null) {
                return look.FirstRow.Value;
            }
            return (WordTable.GetTableLookMask(look) & TableLookFirstRowMask) != 0;
        }

        private static string? NormalizeStyleColor(string? value) {
            if (string.IsNullOrWhiteSpace(value) || string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase)) {
                return null;
            }
            string trimmed = value!.Trim().TrimStart('#');
            return trimmed.Length == 6 && trimmed.All(Uri.IsHexDigit) ? trimmed.ToUpperInvariant() : null;
        }

    }
}
