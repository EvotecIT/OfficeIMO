using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private readonly record struct NativeTableStyleDefaults(PdfCore.PdfCellPadding? CellPadding, PdfCore.PdfColor? CellFill, (PdfCore.PdfColor Color, double Width)? TableBorder, W.TableWidth? PreferredWidth, double? LeftIndent, double? CellSpacing, double? ParagraphLineHeight, double? ParagraphLineSpacingPoints, W.LineSpacingRuleValues? ParagraphLineSpacingRule, double? ParagraphSpacingBefore, double? ParagraphSpacingAfter, NativeTableRunStyleDefaults RunStyle) {
            public static NativeTableStyleDefaults Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null, NativeTableRunStyleDefaults.Empty);
        }

        private readonly record struct NativeTableRunStyleDefaults(double? FontSize, string? FontFamily, bool? Bold, bool? Italic, bool? Underline, bool? Strike, string? ColorHex, W.HighlightColorValues? Highlight) {
            public static NativeTableRunStyleDefaults Empty { get; } = new(null, null, null, null, null, null, null, null);
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
            W.TableWidth? preferredWidth = null;
            double? leftIndent = null;
            double? cellSpacing = null;
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
                leftIndent = GetNativeTableLeftIndent(tableProperties?.GetFirstChild<W.TableIndentation>()) ?? leftIndent;
                cellSpacing = GetNativeTableCellSpacing(tableProperties?.GetFirstChild<W.TableCellSpacing>()) ?? cellSpacing;

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

                W.TableBorders? tableBorders = tableProperties?.GetFirstChild<W.TableBorders>();
                if (tableBorders != null) {
                    tableBorder = GetNativeUniformTableBorder(tableBorders);
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
                preferredWidth,
                leftIndent,
                cellSpacing,
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
                    highlight));
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
