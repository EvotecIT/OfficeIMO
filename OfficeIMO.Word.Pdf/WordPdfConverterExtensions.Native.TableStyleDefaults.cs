using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private readonly record struct NativeTableStyleDefaults(PdfCore.PdfCellPadding? CellPadding, double? ParagraphLineHeight, double? ParagraphLineSpacingPoints, double? ParagraphSpacingAfter) {
            public static NativeTableStyleDefaults Empty { get; } = new(null, null, null, null);
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
            double? paragraphLineHeight = null;
            double? paragraphLineSpacingPoints = null;
            double? paragraphSpacingAfter = null;

            foreach (W.Style style in styleChain) {
                W.TableCellMarginDefault? margins = style.GetFirstChild<W.StyleTableProperties>()?.GetFirstChild<W.TableCellMarginDefault>();
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

                W.SpacingBetweenLines? spacing = style.GetFirstChild<W.StyleParagraphProperties>()?.GetFirstChild<W.SpacingBetweenLines>();
                if (spacing != null) {
                    double? styleParagraphLineHeight = GetNativeTableStyleParagraphLineHeight(spacing);
                    double? styleParagraphLineSpacingPoints = GetNativeTableStyleParagraphLineSpacingPoints(spacing);
                    if (styleParagraphLineHeight.HasValue || styleParagraphLineSpacingPoints.HasValue) {
                        paragraphLineHeight = styleParagraphLineHeight;
                        paragraphLineSpacingPoints = styleParagraphLineSpacingPoints;
                    }

                    paragraphSpacingAfter = ConvertNativeTwipsToPoints(spacing.After?.Value) ?? paragraphSpacingAfter;
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
                paragraphLineHeight,
                paragraphLineSpacingPoints,
                paragraphSpacingAfter);
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
