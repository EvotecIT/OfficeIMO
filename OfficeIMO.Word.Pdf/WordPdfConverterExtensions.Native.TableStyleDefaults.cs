using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private readonly record struct NativeTableStyleDefaults(PdfCore.PdfCellPadding? CellPadding, PdfCore.PdfColor? CellFill, (PdfCore.PdfColor Color, double Width)? TableBorder, W.TableBorders? Borders, W.TableWidth? PreferredWidth, W.TableLayoutValues? Layout, double? LeftIndent, double? CellSpacing, W.TableRowAlignmentValues? Alignment, double? ParagraphLineHeight, double? ParagraphLineSpacingPoints, W.LineSpacingRuleValues? ParagraphLineSpacingRule, double? ParagraphSpacingBefore, double? ParagraphSpacingAfter, NativeTableRunStyleDefaults RunStyle, NativeTableConditionalStyleDefaults FirstRowStyle, NativeTableConditionalStyleDefaults LastRowStyle, NativeTableConditionalStyleDefaults FirstColumnStyle, NativeTableConditionalStyleDefaults LastColumnStyle, NativeTableConditionalStyleDefaults Band1HorizontalStyle, NativeTableConditionalStyleDefaults Band1VerticalStyle) {
            public static NativeTableStyleDefaults Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null, null, null, null, NativeTableRunStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty);
        }

        private readonly record struct NativeTableRunStyleDefaults(double? FontSize, string? FontFamily, bool? Bold, bool? Italic, bool? Underline, bool? Strike, string? ColorHex, W.HighlightColorValues? Highlight, PdfCore.PdfColor? Color) {
            public static NativeTableRunStyleDefaults Empty { get; } = new(null, null, null, null, null, null, null, null, null);
        }

        private readonly record struct NativeTableConditionalStyleDefaults(PdfCore.PdfColor? CellFill, W.TableCellBorders? CellBorders, PdfCore.PdfColor? TextColor, double? FontSize, bool? Bold, bool? Italic, bool? Underline, bool? Strike, W.HighlightColorValues? Highlight) {
            public static NativeTableConditionalStyleDefaults Empty { get; } = new(null, null, null, null, null, null, null, null, null);
        }

        private static NativeTableStyleDefaults GetNativeTableStyleDefaults(WordTable table, NativeDocumentDefaults nativeDefaults, bool ignoreFallbackTableStyle) {
            string? styleId = GetNativeTableStyleId(table);
            if (ignoreFallbackTableStyle && IsNativeFallbackTableStyleId(styleId)) {
                return NativeTableStyleDefaults.Empty;
            }

            IReadOnlyList<W.Style> styleChain = GetNativeTableStyleChain(table.Document, styleId);
            if (styleChain.Count == 0) {
                return NativeTableStyleDefaults.Empty;
            }

            double? marginTop = null;
            double? marginBottom = null;
            double? marginLeft = null;
            double? marginRight = null;
            PdfCore.PdfColor? cellFill = null;
            (PdfCore.PdfColor Color, double Width)? tableBorder = null;
            W.TableBorders? tableBorders = null;
            W.TableWidth? preferredWidth = null;
            W.TableLayoutValues? layout = null;
            double? leftIndent = null;
            double? cellSpacing = null;
            W.TableRowAlignmentValues? alignment = null;
            double? paragraphLineHeight = null;
            double? paragraphLineSpacingPoints = null;
            W.LineSpacingRuleValues? paragraphLineSpacingRule = null;
            double? paragraphSpacingBefore = null;
            double? paragraphSpacingAfter = null;
            double? fontSize = null;
            string? fontFamily = null;
            bool? bold = null;
            bool? italic = null;
            bool? underline = null;
            bool? strike = null;
            string? colorHex = null;
            W.HighlightColorValues? highlight = null;
            NativeTableConditionalStyleDefaults firstRowStyle = NativeTableConditionalStyleDefaults.Empty;
            NativeTableConditionalStyleDefaults lastRowStyle = NativeTableConditionalStyleDefaults.Empty;
            NativeTableConditionalStyleDefaults firstColumnStyle = NativeTableConditionalStyleDefaults.Empty;
            NativeTableConditionalStyleDefaults lastColumnStyle = NativeTableConditionalStyleDefaults.Empty;
            NativeTableConditionalStyleDefaults band1HorizontalStyle = NativeTableConditionalStyleDefaults.Empty;
            NativeTableConditionalStyleDefaults band1VerticalStyle = NativeTableConditionalStyleDefaults.Empty;

            foreach (W.Style style in styleChain) {
                W.StyleRunProperties? runProperties = style.GetFirstChild<W.StyleRunProperties>();
                fontSize = GetNativeStyleFontSize(runProperties) ?? fontSize;
                fontFamily = ResolveNativeRunFontsFamily(table.Document, runProperties?.GetFirstChild<W.RunFonts>()) ?? fontFamily;
                bold = ReadNativeOnOff(runProperties?.GetFirstChild<W.Bold>()) ?? bold;
                italic = ReadNativeOnOff(runProperties?.GetFirstChild<W.Italic>()) ?? italic;
                underline = ReadNativeUnderline(runProperties?.GetFirstChild<W.Underline>()) ?? underline;
                strike = ReadNativeOnOff(runProperties?.GetFirstChild<W.Strike>()) ?? ReadNativeOnOff(runProperties?.GetFirstChild<W.DoubleStrike>()) ?? strike;
                colorHex = runProperties?.GetFirstChild<W.Color>()?.Val?.Value ?? colorHex;
                highlight = runProperties?.GetFirstChild<W.Highlight>()?.Val?.Value ?? highlight;

                W.StyleTableProperties? tableProperties = style.GetFirstChild<W.StyleTableProperties>();
                preferredWidth = tableProperties?.GetFirstChild<W.TableWidth>() ?? preferredWidth;
                layout = tableProperties?.GetFirstChild<W.TableLayout>()?.Type?.Value ?? layout;
                leftIndent = GetNativeTableLeftIndent(tableProperties?.GetFirstChild<W.TableIndentation>()) ?? leftIndent;
                cellSpacing = GetNativeTableCellSpacing(tableProperties?.GetFirstChild<W.TableCellSpacing>()) ?? cellSpacing;
                alignment = tableProperties?.GetFirstChild<W.TableJustification>()?.Val?.Value ?? alignment;

                W.TableCellMarginDefault? margins = tableProperties?.GetFirstChild<W.TableCellMarginDefault>();
                if (margins != null) {
                    double? top = ConvertNativeTwipsToPoints(margins.TopMargin?.Width?.Value);
                    double? bottom = ConvertNativeTwipsToPoints(margins.BottomMargin?.Width?.Value);
                    double? left = margins.TableCellLeftMargin?.Width == null
                        ? null
                        : ConvertNativeTwipsToPoints(margins.TableCellLeftMargin.Width.Value);
                    double? right = margins.TableCellRightMargin?.Width == null
                        ? null
                        : ConvertNativeTwipsToPoints(margins.TableCellRightMargin.Width.Value);

                    marginTop = top ?? marginTop;
                    marginBottom = bottom ?? marginBottom;
                    marginLeft = left ?? marginLeft;
                    marginRight = right ?? marginRight;
                }

                W.Shading? shading = tableProperties?.GetFirstChild<W.Shading>();
                if (shading != null) {
                    cellFill = ParseNativeColor(shading.Fill?.Value);
                }

                W.TableBorders? styleTableBorders = tableProperties?.GetFirstChild<W.TableBorders>();
                if (styleTableBorders != null) {
                    tableBorder = GetNativeUniformTableBorder(styleTableBorders);
                    tableBorders = styleTableBorders;
                }

                W.SpacingBetweenLines? spacing = style.GetFirstChild<W.StyleParagraphProperties>()?.GetFirstChild<W.SpacingBetweenLines>();
                if (spacing != null) {
                    double? styleParagraphLineHeight = GetNativeTableStyleParagraphLineHeight(spacing);
                    double? styleParagraphLineSpacingPoints = GetNativeTableStyleParagraphLineSpacingPoints(spacing);
                    if (styleParagraphLineHeight.HasValue || styleParagraphLineSpacingPoints.HasValue) {
                        paragraphLineHeight = styleParagraphLineHeight;
                        paragraphLineSpacingPoints = styleParagraphLineSpacingPoints;
                        paragraphLineSpacingRule = spacing.LineRule?.Value;
                    }

                    double effectiveFontSize = fontSize ?? nativeDefaults.FontSize;
                    double effectiveLineHeight = styleParagraphLineSpacingPoints.HasValue && effectiveFontSize > 0D
                        ? ResolveNativeLineSpacingHeight(styleParagraphLineSpacingPoints.Value, spacing.LineRule?.Value, effectiveFontSize, NativeWordTableSingleLineHeight)
                        : styleParagraphLineHeight ?? paragraphLineHeight ?? NativeWordTableSingleLineHeight;
                    paragraphSpacingBefore = GetNativeSpacingBeforePoints(spacing, effectiveFontSize, effectiveLineHeight) ?? paragraphSpacingBefore;
                    paragraphSpacingAfter = GetNativeSpacingAfterPoints(spacing, effectiveFontSize, effectiveLineHeight) ?? paragraphSpacingAfter;
                }

                firstRowStyle = GetNativeTableConditionalStyleDefaults(style, W.TableStyleOverrideValues.FirstRow, firstRowStyle);
                lastRowStyle = GetNativeTableConditionalStyleDefaults(style, W.TableStyleOverrideValues.LastRow, lastRowStyle);
                firstColumnStyle = GetNativeTableConditionalStyleDefaults(style, W.TableStyleOverrideValues.FirstColumn, firstColumnStyle);
                lastColumnStyle = GetNativeTableConditionalStyleDefaults(style, W.TableStyleOverrideValues.LastColumn, lastColumnStyle);
                band1HorizontalStyle = GetNativeTableConditionalStyleDefaults(style, W.TableStyleOverrideValues.Band1Horizontal, band1HorizontalStyle);
                band1VerticalStyle = GetNativeTableConditionalStyleDefaults(style, W.TableStyleOverrideValues.Band1Vertical, band1VerticalStyle);
            }

            PdfCore.PdfCellPadding? cellPadding = marginTop.HasValue || marginBottom.HasValue || marginLeft.HasValue || marginRight.HasValue
                ? new PdfCore.PdfCellPadding {
                    Top = marginTop,
                    Bottom = marginBottom,
                    Left = marginLeft,
                    Right = marginRight
                }
                : null;

            return new NativeTableStyleDefaults(
                cellPadding,
                cellFill,
                tableBorder,
                tableBorders,
                preferredWidth,
                layout,
                leftIndent,
                cellSpacing,
                alignment,
                paragraphLineHeight,
                paragraphLineSpacingPoints,
                paragraphLineSpacingRule,
                paragraphSpacingBefore,
                paragraphSpacingAfter,
                new NativeTableRunStyleDefaults(
                    fontSize,
                    fontFamily,
                    bold,
                    italic,
                    underline,
                    strike,
                    colorHex,
                    highlight,
                    null),
                firstRowStyle,
                lastRowStyle,
                firstColumnStyle,
                lastColumnStyle,
                band1HorizontalStyle,
                band1VerticalStyle);
        }

        private static NativeTableConditionalStyleDefaults GetNativeTableConditionalStyleDefaults(W.Style style, W.TableStyleOverrideValues type, NativeTableConditionalStyleDefaults inherited) {
            NativeTableConditionalStyleDefaults result = inherited;
            foreach (W.TableStyleProperties properties in style.Elements<W.TableStyleProperties>().Where(properties => properties.Type?.Value == type)) {
                W.TableStyleConditionalFormattingTableCellProperties? cellProperties = properties.GetFirstChild<W.TableStyleConditionalFormattingTableCellProperties>();
                PdfCore.PdfColor? cellFill = ParseNativeColor(cellProperties?.GetFirstChild<W.Shading>()?.Fill?.Value);
                W.TableCellBorders? cellBorders = cellProperties?.GetFirstChild<W.TableCellBorders>();

                W.RunPropertiesBaseStyle? runProperties = properties.GetFirstChild<W.RunPropertiesBaseStyle>();
                PdfCore.PdfColor? textColor = ParseNativeColor(runProperties?.GetFirstChild<W.Color>()?.Val?.Value);
                double? fontSize = GetNativeRunPropertiesBaseStyleFontSize(runProperties);
                bool? bold = ReadNativeOnOff(runProperties?.GetFirstChild<W.Bold>());
                bool? italic = ReadNativeOnOff(runProperties?.GetFirstChild<W.Italic>());
                bool? underline = ReadNativeUnderline(runProperties?.GetFirstChild<W.Underline>());
                bool? strike = ReadNativeOnOff(runProperties?.GetFirstChild<W.Strike>()) ?? ReadNativeOnOff(runProperties?.GetFirstChild<W.DoubleStrike>());
                W.HighlightColorValues? highlight = runProperties?.GetFirstChild<W.Highlight>()?.Val?.Value;
                result = new NativeTableConditionalStyleDefaults(
                    cellFill ?? result.CellFill,
                    cellBorders ?? result.CellBorders,
                    textColor ?? result.TextColor,
                    fontSize ?? result.FontSize,
                    bold ?? result.Bold,
                    italic ?? result.Italic,
                    underline ?? result.Underline,
                    strike ?? result.Strike,
                    highlight ?? result.Highlight);
            }

            return result;
        }

        private static double? GetNativeRunPropertiesBaseStyleFontSize(W.RunPropertiesBaseStyle? runProperties) {
            string? value = runProperties?.FontSize?.Val?.Value;
            if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double halfPoints) ||
                halfPoints <= 0D ||
                double.IsNaN(halfPoints) ||
                double.IsInfinity(halfPoints)) {
                return null;
            }

            return halfPoints / 2D;
        }

        private static double? GetNativeTableStyleParagraphLineHeight(W.SpacingBetweenLines spacing) {
            if (spacing.LineRule?.Value != W.LineSpacingRuleValues.Auto) {
                return null;
            }

            if (!double.TryParse(spacing.Line?.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double line) ||
                line <= 0D ||
                double.IsNaN(line) ||
                double.IsInfinity(line)) {
                return null;
            }

            return Math.Max(0.01D, NativeWordTableSingleLineHeight * (line / 240D));
        }

        private static double? GetNativeTableStyleParagraphLineSpacingPoints(W.SpacingBetweenLines spacing) {
            if (spacing.LineRule?.Value == W.LineSpacingRuleValues.Auto) {
                return null;
            }

            return ConvertNativeTwipsToPoints(spacing.Line?.Value);
        }

        private static IReadOnlyList<W.Style> GetNativeTableStyleChain(WordDocument document, string? styleId) {
            W.Styles? styles = document._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            if (styles == null) {
                return Array.Empty<W.Style>();
            }

            Dictionary<string, W.Style> tableStyles = styles
                .Elements<W.Style>()
                .Where(style => style.Type?.Value == W.StyleValues.Table && !string.IsNullOrEmpty(style.StyleId?.Value))
                .ToDictionary(style => style.StyleId!.Value!, style => style, StringComparer.Ordinal);

            if (string.IsNullOrWhiteSpace(styleId)) {
                styleId = tableStyles.Values.FirstOrDefault(style => style.Default?.Value == true)?.StyleId?.Value;
            }

            if (string.IsNullOrWhiteSpace(styleId)) {
                return Array.Empty<W.Style>();
            }

            var chain = new List<W.Style>();
            var visited = new HashSet<string>(StringComparer.Ordinal);
            string? currentStyleId = styleId;
            while (!string.IsNullOrWhiteSpace(currentStyleId) && visited.Add(currentStyleId!) && tableStyles.TryGetValue(currentStyleId!, out W.Style? style)) {
                chain.Add(style);
                currentStyleId = style.BasedOn?.Val?.Value;
            }

            chain.Reverse();
            return chain;
        }

        private static string? GetNativeTableStyleId(WordTable table) =>
            table._tableProperties?.TableStyle?.Val?.Value;

        private static bool IsNativeFallbackTableStyleId(string? styleId) =>
            string.IsNullOrWhiteSpace(styleId) ||
            string.Equals(styleId, "TableNormal", StringComparison.Ordinal);
    }
}
