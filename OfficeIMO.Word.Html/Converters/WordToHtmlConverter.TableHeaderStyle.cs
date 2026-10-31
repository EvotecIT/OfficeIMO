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
            internal string? Fill { get; set; }
            internal string? TopBorder { get; set; }
            internal string? BottomBorder { get; set; }

            internal IEnumerable<string> ToCss(bool hasDirectFill, bool hasDirectTopBorder, bool hasDirectBottomBorder) {
                if (Bold) yield return "font-weight:bold";
                if (Color != null) yield return "color:#" + Color;
                if (Fill != null && !hasDirectFill) yield return "background-color:#" + Fill;
                if (TopBorder != null && !hasDirectTopBorder) yield return "border-top:" + TopBorder;
                if (BottomBorder != null && !hasDirectBottomBorder) yield return "border-bottom:" + BottomBorder;
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
            foreach (TableStyleProperties conditional in conditionals) {
                RunPropertiesBaseStyle? run = conditional.GetFirstChild<RunPropertiesBaseStyle>();
                if (bold == null && run?.GetFirstChild<Bold>() is Bold boldElement) {
                    bold = boldElement.Val == null || boldElement.Val.Value;
                }
                appearance.Color ??= NormalizeStyleColor(run?.GetFirstChild<Color>()?.Val?.Value);
                var cell = conditional.GetFirstChild<TableStyleConditionalFormattingTableCellProperties>();
                appearance.Fill ??= NormalizeStyleColor(cell?.GetFirstChild<Shading>()?.Fill?.Value);
                TableCellBorders? borders = cell?.GetFirstChild<TableCellBorders>();
                appearance.TopBorder ??= FormatStyleBorderCss(borders?.TopBorder);
                appearance.BottomBorder ??= FormatStyleBorderCss(borders?.BottomBorder);
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

        private static string? FormatStyleBorderCss(BorderType? border) {
            if (border?.Val == null) {
                return null;
            }
            BorderValues value = border.Val.Value;
            if (value == BorderValues.Nil || value == BorderValues.None) {
                return null;
            }
            string style = value == BorderValues.Double ? "double"
                : value == BorderValues.Dotted ? "dotted"
                : value == BorderValues.Dashed || value == BorderValues.DashSmallGap ? "dashed"
                : "solid";
            double widthPoints = Math.Max(0.5D, (border.Size?.Value ?? 4U) / 8D);
            if (style == "double") {
                widthPoints = Math.Max(widthPoints, 2.25D);
            }
            string color = NormalizeStyleColor(border.Color?.Value) ?? "000000";
            return Math.Round(widthPoints, 2).ToString("0.##", CultureInfo.InvariantCulture) + "pt " + style + " #" + color;
        }
    }
}
