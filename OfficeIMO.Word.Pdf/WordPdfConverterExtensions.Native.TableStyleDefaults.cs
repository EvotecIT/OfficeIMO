using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private readonly record struct NativeTableStyleDefaults(PdfCore.PdfCellPadding? CellPadding, PdfCore.PdfColor? CellFill, PdfCore.PdfCellVerticalAlign? CellVerticalAlignment, (PdfCore.PdfColor Color, double Width)? TableBorder, W.TableBorders? Borders, W.TableWidth? PreferredWidth, W.TableLayoutValues? Layout, double? LeftIndent, double? CellSpacing, W.TableRowAlignmentValues? Alignment, double? ParagraphLineHeight, double? ParagraphLineSpacingPoints, W.LineSpacingRuleValues? ParagraphLineSpacingRule, double? ParagraphSpacingBefore, double? ParagraphSpacingAfter, W.JustificationValues? ParagraphAlignment, double? ParagraphLeftIndent, double? ParagraphRightIndent, double? ParagraphFirstLineIndent, NativeTableRunStyleDefaults RunStyle, NativeTableConditionalStyleDefaults FirstRowStyle, NativeTableConditionalStyleDefaults LastRowStyle, NativeTableConditionalStyleDefaults FirstColumnStyle, NativeTableConditionalStyleDefaults LastColumnStyle, NativeTableConditionalStyleDefaults Band1HorizontalStyle, NativeTableConditionalStyleDefaults Band1VerticalStyle) {
            public static NativeTableStyleDefaults Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, NativeTableRunStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty);
        }

        private readonly record struct NativeTableRunStyleDefaults(double? FontSize, string? FontFamily, bool? Bold, bool? Italic, bool? Underline, bool? Strike, bool? AllCaps, W.VerticalPositionValues? Baseline, string? ColorHex, W.HighlightColorValues? Highlight, PdfCore.PdfColor? Color) {
            public static NativeTableRunStyleDefaults Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null);
        }

        private readonly record struct NativeTableConditionalStyleDefaults(PdfCore.PdfColor? CellFill, W.TableCellBorders? CellBorders, PdfCore.PdfCellPadding? CellPadding, PdfCore.PdfCellVerticalAlign? CellVerticalAlignment, PdfCore.PdfColor? TextColor, double? FontSize, bool? Bold, bool? Italic, bool? Underline, bool? Strike, bool? AllCaps, W.VerticalPositionValues? Baseline, W.HighlightColorValues? Highlight, double? ParagraphLineHeight, double? ParagraphLineSpacingPoints, W.LineSpacingRuleValues? ParagraphLineSpacingRule, double? ParagraphSpacingBefore, double? ParagraphSpacingAfter, W.JustificationValues? ParagraphAlignment, double? ParagraphLeftIndent, double? ParagraphRightIndent, double? ParagraphFirstLineIndent) {
            public static NativeTableConditionalStyleDefaults Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
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
            PdfCore.PdfCellVerticalAlign? cellVerticalAlignment = null;
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
            W.JustificationValues? paragraphAlignment = null;
            double? paragraphLeftIndent = null;
            double? paragraphRightIndent = null;
            double? paragraphFirstLineIndent = null;
            double? fontSize = null;
            string? fontFamily = null;
            bool? bold = null;
            bool? italic = null;
            bool? underline = null;
            bool? strike = null;
            bool? allCaps = null;
            W.VerticalPositionValues? baseline = null;
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
                allCaps = ReadNativeOnOff(runProperties?.GetFirstChild<W.Caps>()) ?? ReadNativeOnOff(runProperties?.GetFirstChild<W.SmallCaps>()) ?? allCaps;
                baseline = runProperties?.GetFirstChild<W.VerticalTextAlignment>()?.Val?.Value ?? baseline;
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

                W.StyleTableCellProperties? tableCellProperties = style.GetFirstChild<W.StyleTableCellProperties>();
                cellVerticalAlignment = MapNativeNullableCellVerticalAlign(tableCellProperties?.GetFirstChild<W.TableCellVerticalAlignment>()?.Val?.Value) ?? cellVerticalAlignment;

                W.TableBorders? styleTableBorders = tableProperties?.GetFirstChild<W.TableBorders>();
                if (styleTableBorders != null) {
                    tableBorder = GetNativeUniformTableBorder(styleTableBorders);
                    tableBorders = styleTableBorders;
                }

                W.StyleParagraphProperties? paragraphProperties = style.GetFirstChild<W.StyleParagraphProperties>();
                W.Indentation? indentation = paragraphProperties?.GetFirstChild<W.Indentation>();
                if (indentation != null) {
                    paragraphLeftIndent = ConvertNativeTwipsToPoints(indentation.Left?.Value) ?? paragraphLeftIndent;
                    paragraphRightIndent = ConvertNativeTwipsToPoints(indentation.Right?.Value) ?? paragraphRightIndent;

                    double? firstLine = ConvertNativeTwipsToPoints(indentation.FirstLine?.Value);
                    double? hanging = ConvertNativeTwipsToPoints(indentation.Hanging?.Value);
                    if (hanging.HasValue) {
                        paragraphFirstLineIndent = -hanging.Value;
                    } else if (firstLine.HasValue) {
                        paragraphFirstLineIndent = firstLine.Value;
                    }
                }

                W.SpacingBetweenLines? spacing = paragraphProperties?.GetFirstChild<W.SpacingBetweenLines>();
                if (spacing != null) {
                    double naturalLineHeight = ResolveNativeWordSingleLineHeight(fontFamily, nativeDefaults.FontFamily);
                    double? styleParagraphLineHeight = GetNativeTableStyleParagraphLineHeight(spacing, fontFamily, nativeDefaults.FontFamily);
                    double? styleParagraphLineSpacingPoints = GetNativeTableStyleParagraphLineSpacingPoints(spacing);
                    if (styleParagraphLineHeight.HasValue || styleParagraphLineSpacingPoints.HasValue) {
                        paragraphLineHeight = styleParagraphLineHeight;
                        paragraphLineSpacingPoints = styleParagraphLineSpacingPoints;
                        paragraphLineSpacingRule = spacing.LineRule?.Value;
                    }

                    double effectiveFontSize = fontSize ?? nativeDefaults.FontSize;
                    double effectiveLineHeight = styleParagraphLineSpacingPoints.HasValue && effectiveFontSize > 0D
                        ? ResolveNativeLineSpacingHeight(styleParagraphLineSpacingPoints.Value, spacing.LineRule?.Value, effectiveFontSize, naturalLineHeight)
                        : styleParagraphLineHeight ?? paragraphLineHeight ?? naturalLineHeight;
                    paragraphSpacingBefore = GetNativeSpacingBeforePoints(spacing, effectiveFontSize, effectiveLineHeight) ?? paragraphSpacingBefore;
                    paragraphSpacingAfter = GetNativeSpacingAfterPoints(spacing, effectiveFontSize, effectiveLineHeight) ?? paragraphSpacingAfter;
                }

                paragraphAlignment = paragraphProperties?.GetFirstChild<W.Justification>()?.Val?.Value ?? paragraphAlignment;

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
                cellVerticalAlignment,
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
                paragraphAlignment,
                paragraphLeftIndent,
                paragraphRightIndent,
                paragraphFirstLineIndent,
                new NativeTableRunStyleDefaults(
                    fontSize,
                    fontFamily,
                    bold,
                    italic,
                    underline,
                    strike,
                    allCaps,
                    baseline,
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
                PdfCore.PdfCellPadding? cellPadding = CreateNativeConditionalTableCellPadding(cellProperties?.GetFirstChild<W.TableCellMargin>());
                PdfCore.PdfCellVerticalAlign? cellVerticalAlignment = MapNativeNullableCellVerticalAlign(cellProperties?.GetFirstChild<W.TableCellVerticalAlignment>()?.Val?.Value);

                W.RunPropertiesBaseStyle? runProperties = properties.GetFirstChild<W.RunPropertiesBaseStyle>();
                W.RunFonts? runFonts = runProperties?.GetFirstChild<W.RunFonts>();
                string? fontFamily = FirstNonWhiteSpace(
                    runFonts?.Ascii?.Value,
                    runFonts?.HighAnsi?.Value,
                    runFonts?.EastAsia?.Value,
                    runFonts?.ComplexScript?.Value);
                PdfCore.PdfColor? textColor = ParseNativeColor(runProperties?.GetFirstChild<W.Color>()?.Val?.Value);
                double? fontSize = GetNativeRunPropertiesBaseStyleFontSize(runProperties);
                bool? bold = ReadNativeOnOff(runProperties?.GetFirstChild<W.Bold>());
                bool? italic = ReadNativeOnOff(runProperties?.GetFirstChild<W.Italic>());
                bool? underline = ReadNativeUnderline(runProperties?.GetFirstChild<W.Underline>());
                bool? strike = ReadNativeOnOff(runProperties?.GetFirstChild<W.Strike>()) ?? ReadNativeOnOff(runProperties?.GetFirstChild<W.DoubleStrike>());
                bool? allCaps = ReadNativeOnOff(runProperties?.GetFirstChild<W.Caps>()) ?? ReadNativeOnOff(runProperties?.GetFirstChild<W.SmallCaps>());
                W.VerticalPositionValues? baseline = runProperties?.GetFirstChild<W.VerticalTextAlignment>()?.Val?.Value;
                W.HighlightColorValues? highlight = runProperties?.GetFirstChild<W.Highlight>()?.Val?.Value;

                W.StyleParagraphProperties? paragraphProperties = properties.GetFirstChild<W.StyleParagraphProperties>();
                W.SpacingBetweenLines? spacing = paragraphProperties?.GetFirstChild<W.SpacingBetweenLines>();
                double? paragraphLineHeight = null;
                double? paragraphLineSpacingPoints = null;
                W.LineSpacingRuleValues? paragraphLineSpacingRule = null;
                double? paragraphSpacingBefore = null;
                double? paragraphSpacingAfter = null;
                double? paragraphLeftIndent = null;
                double? paragraphRightIndent = null;
                double? paragraphFirstLineIndent = null;
                W.Indentation? indentation = paragraphProperties?.GetFirstChild<W.Indentation>();
                if (indentation != null) {
                    paragraphLeftIndent = ConvertNativeTwipsToPoints(indentation.Left?.Value);
                    paragraphRightIndent = ConvertNativeTwipsToPoints(indentation.Right?.Value);

                    double? firstLine = ConvertNativeTwipsToPoints(indentation.FirstLine?.Value);
                    double? hanging = ConvertNativeTwipsToPoints(indentation.Hanging?.Value);
                    if (hanging.HasValue) {
                        paragraphFirstLineIndent = -hanging.Value;
                    } else if (firstLine.HasValue) {
                        paragraphFirstLineIndent = firstLine.Value;
                    }
                }

                if (spacing != null) {
                    double naturalLineHeight = ResolveNativeWordSingleLineHeight(fontFamily);
                    paragraphLineHeight = GetNativeTableStyleParagraphLineHeight(spacing, fontFamily);
                    paragraphLineSpacingPoints = GetNativeTableStyleParagraphLineSpacingPoints(spacing);
                    paragraphLineSpacingRule = paragraphLineHeight.HasValue || paragraphLineSpacingPoints.HasValue
                        ? spacing.LineRule?.Value
                        : null;

                    double effectiveFontSize = fontSize ?? result.FontSize ?? NativeDocumentDefaults.WordDefault.FontSize;
                    double effectiveLineHeight = paragraphLineSpacingPoints.HasValue && effectiveFontSize > 0D
                        ? ResolveNativeLineSpacingHeight(paragraphLineSpacingPoints.Value, spacing.LineRule?.Value, effectiveFontSize, naturalLineHeight)
                        : paragraphLineHeight ?? result.ParagraphLineHeight ?? naturalLineHeight;
                    paragraphSpacingBefore = GetNativeSpacingBeforePoints(spacing, effectiveFontSize, effectiveLineHeight);
                    paragraphSpacingAfter = GetNativeSpacingAfterPoints(spacing, effectiveFontSize, effectiveLineHeight);
                }

                W.JustificationValues? paragraphAlignment = paragraphProperties?.GetFirstChild<W.Justification>()?.Val?.Value;
                result = new NativeTableConditionalStyleDefaults(
                    cellFill ?? result.CellFill,
                    cellBorders ?? result.CellBorders,
                    MergeNativeCellPadding(result.CellPadding, cellPadding),
                    cellVerticalAlignment ?? result.CellVerticalAlignment,
                    textColor ?? result.TextColor,
                    fontSize ?? result.FontSize,
                    bold ?? result.Bold,
                    italic ?? result.Italic,
                    underline ?? result.Underline,
                    strike ?? result.Strike,
                    allCaps ?? result.AllCaps,
                    baseline ?? result.Baseline,
                    highlight ?? result.Highlight,
                    paragraphLineHeight ?? result.ParagraphLineHeight,
                    paragraphLineSpacingPoints ?? result.ParagraphLineSpacingPoints,
                    paragraphLineSpacingRule ?? result.ParagraphLineSpacingRule,
                    paragraphSpacingBefore ?? result.ParagraphSpacingBefore,
                    paragraphSpacingAfter ?? result.ParagraphSpacingAfter,
                    paragraphAlignment ?? result.ParagraphAlignment,
                    paragraphLeftIndent ?? result.ParagraphLeftIndent,
                    paragraphRightIndent ?? result.ParagraphRightIndent,
                    paragraphFirstLineIndent ?? result.ParagraphFirstLineIndent);
            }

            return result;
        }

        private static PdfCore.PdfCellPadding? CreateNativeConditionalTableCellPadding(W.TableCellMargin? margins) {
            if (margins == null) {
                return null;
            }

            double? top = ConvertNativeTwipsToPoints(margins.TopMargin?.Width?.Value);
            double? bottom = ConvertNativeTwipsToPoints(margins.BottomMargin?.Width?.Value);
            double? left = ConvertNativeTwipsToPoints(margins.LeftMargin?.Width?.Value);
            double? right = ConvertNativeTwipsToPoints(margins.RightMargin?.Width?.Value);
            if (!top.HasValue && !bottom.HasValue && !left.HasValue && !right.HasValue) {
                return null;
            }

            return new PdfCore.PdfCellPadding {
                Top = top,
                Bottom = bottom,
                Left = left,
                Right = right
            };
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

        private static double? GetNativeTableStyleParagraphLineHeight(
            W.SpacingBetweenLines spacing,
            params string?[] fontFamilies) {
            if (spacing.LineRule?.Value != W.LineSpacingRuleValues.Auto) {
                return null;
            }

            if (!double.TryParse(spacing.Line?.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double line) ||
                line <= 0D ||
                double.IsNaN(line) ||
                double.IsInfinity(line)) {
                return null;
            }

            return Math.Max(0.01D, ResolveNativeWordSingleLineHeight(fontFamilies) * (line / 240D));
        }

        private static double? GetNativeTableStyleParagraphLineSpacingPoints(W.SpacingBetweenLines spacing) {
            if (spacing.LineRule?.Value == W.LineSpacingRuleValues.Auto) {
                return null;
            }

            return ConvertNativeTwipsToPoints(spacing.Line?.Value);
        }

        private static IReadOnlyList<W.Style> GetNativeTableStyleChain(WordDocument document, string? styleId) {
            NativeStyleLookupCache? cache = GetNativeStyleLookupCache(document);
            string? resolvedStyleId = cache?.ResolveTableStyleId(styleId);
            if (cache == null || string.IsNullOrWhiteSpace(resolvedStyleId)) {
                return Array.Empty<W.Style>();
            }

            if (cache.TableChains.TryGetValue(resolvedStyleId!, out IReadOnlyList<W.Style>? cachedChain)) {
                return cachedChain;
            }

            var chain = new List<W.Style>();
            var visited = new HashSet<string>(StringComparer.Ordinal);
            string? currentStyleId = resolvedStyleId;
            while (!string.IsNullOrWhiteSpace(currentStyleId) && visited.Add(currentStyleId!) && cache.TableStyles.TryGetValue(currentStyleId!, out W.Style? style)) {
                cache.RecordStyleChainReference(chain.Count + 1);
                chain.Add(style);
                currentStyleId = style.BasedOn?.Val?.Value;
            }

            chain.Reverse();
            IReadOnlyList<W.Style> result = chain.ToArray();
            cache.TableChains[resolvedStyleId!] = result;
            return result;
        }

        private static string? GetNativeTableStyleId(WordTable table) =>
            table._tableProperties?.TableStyle?.Val?.Value;

        private static bool IsNativeFallbackTableStyleId(string? styleId) =>
            string.IsNullOrWhiteSpace(styleId) ||
            string.Equals(styleId, "TableNormal", StringComparison.Ordinal);
    }
}
