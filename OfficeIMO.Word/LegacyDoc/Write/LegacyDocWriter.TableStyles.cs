using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static IReadOnlyDictionary<string, Style> ReadTableStyleDefinitions(MainDocumentPart mainPart) {
            Styles? styles = mainPart.StyleDefinitionsPart?.Styles;
            if (styles == null) {
                return new Dictionary<string, Style>(StringComparer.OrdinalIgnoreCase);
            }

            return styles
                .Elements<Style>()
                .Where(style => style.Type?.Value == StyleValues.Table)
                .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
                .GroupBy(style => style.StyleId!.Value!, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
        }

        private static LegacyDocTableBorders ReadSupportedTableStyleBorders(TableStyle? tableStyle, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            string? styleId = tableStyle?.Val?.Value;
            if (IsNoOpTableStyle(styleId)) {
                return default;
            }

            if (IsTableGridStyle(styleId)) {
                return ReadSupportedTableGridBorders();
            }

            if (string.IsNullOrWhiteSpace(styleId) || !tableStyleDefinitions.TryGetValue(styleId!, out Style? style)) {
                throw new NotSupportedException($"Native DOC saving supports simple tables only when table style '{styleId}' can be resolved to supported table-level formatting.");
            }

            ThrowIfUnsupportedTableStyle(styleId!, style, tableStyleDefinitions);
            LegacyDocTableBorders inheritedBorders = ReadSupportedTableStyleBaseBorders(style, tableStyleDefinitions);
            TableBorders? customBorders = style.GetFirstChild<StyleTableProperties>()?.GetFirstChild<TableBorders>();
            LegacyDocTableBorders ownBorders = customBorders == null ? default : ReadSupportedTableBorders(customBorders);
            return MergeSupportedTableBorders(ownBorders, inheritedBorders);
        }

        private static LegacyDocTableCellShading ReadSupportedTableStyleShading(TableStyle? tableStyle, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            string? styleId = tableStyle?.Val?.Value;
            if (IsNoOpTableStyle(styleId)) {
                return default;
            }

            if (IsTableGridStyle(styleId)) {
                Style styleDefinition = WordTableStyles.GetStyleDefinition(WordTableStyle.TableGrid);
                Shading? shading = styleDefinition.GetFirstChild<StyleTableProperties>()?.GetFirstChild<Shading>();
                return shading == null ? default : ReadSupportedTableCellShading(shading, "table style shading");
            }

            if (string.IsNullOrWhiteSpace(styleId) || !tableStyleDefinitions.TryGetValue(styleId!, out Style? style)) {
                throw new NotSupportedException($"Native DOC saving supports simple tables only when table style '{styleId}' can be resolved to supported table-level formatting.");
            }

            ThrowIfUnsupportedTableStyle(styleId!, style, tableStyleDefinitions);
            LegacyDocTableCellShading inheritedShading = ReadSupportedTableStyleBaseShading(style, tableStyleDefinitions);
            Shading? customShading = style.GetFirstChild<StyleTableProperties>()?.GetFirstChild<Shading>();
            LegacyDocTableCellShading ownShading = customShading == null ? default : ReadSupportedTableCellShading(customShading, "table style shading");
            return ownShading.HasAny ? ownShading : inheritedShading;
        }

        private static LegacyDocTableCellMargins? ReadSupportedTableStyleDefaultCellMargins(TableStyle? tableStyle, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            Style? style = ResolveSupportedTableStyle(tableStyle, tableStyleDefinitions);
            TableCellMarginDefault? margins = style?.GetFirstChild<StyleTableProperties>()?.GetFirstChild<TableCellMarginDefault>();
            return margins == null ? null : ReadSupportedTableDefaultCellMargins(margins);
        }

        private static int? ReadSupportedTableStyleDefaultCellSpacing(TableStyle? tableStyle, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            Style? style = ResolveSupportedTableStyle(tableStyle, tableStyleDefinitions);
            TableCellSpacing? spacing = style?.GetFirstChild<StyleTableProperties>()?.GetFirstChild<TableCellSpacing>();
            return spacing == null ? null : ReadSupportedTableDefaultCellSpacing(spacing);
        }

        private static LegacyDocTableAlignment? ReadSupportedTableStyleAlignment(TableStyle? tableStyle, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            Style? style = ResolveSupportedTableStyle(tableStyle, tableStyleDefinitions);
            TableJustification? justification = style?.GetFirstChild<StyleTableProperties>()?.GetFirstChild<TableJustification>();
            return justification == null ? null : ReadSupportedTableAlignment(justification);
        }

        private static int? ReadSupportedTableStyleIndentation(TableStyle? tableStyle, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            Style? style = ResolveSupportedTableStyle(tableStyle, tableStyleDefinitions);
            TableIndentation? indentation = style?.GetFirstChild<StyleTableProperties>()?.GetFirstChild<TableIndentation>();
            return indentation == null ? null : ReadSupportedTableIndentation(indentation);
        }

        private static LegacyDocTablePreferredWidth? ReadSupportedTableStylePreferredWidth(TableStyle? tableStyle, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            Style? style = ResolveSupportedTableStyle(tableStyle, tableStyleDefinitions);
            TableWidth? width = style?.GetFirstChild<StyleTableProperties>()?.GetFirstChild<TableWidth>();
            return width == null ? null : ReadSupportedTablePreferredWidth(width);
        }

        private static bool? ReadSupportedTableStyleAutofit(TableStyle? tableStyle, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            Style? style = ResolveSupportedTableStyle(tableStyle, tableStyleDefinitions);
            TableLayout? layout = style?.GetFirstChild<StyleTableProperties>()?.GetFirstChild<TableLayout>();
            return layout == null ? null : ReadSupportedTableAutofit(layout);
        }

        private static LegacyDocWritableParagraphFormatting ReadSupportedTableStyleParagraphFormatting(TableStyle? tableStyle, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            string? styleId = tableStyle?.Val?.Value;
            if (IsNoOpTableStyle(styleId) || string.Equals(styleId, "TableGrid", StringComparison.OrdinalIgnoreCase)) {
                return LegacyDocWritableParagraphFormatting.Plain;
            }

            Style? style = ResolveSupportedTableStyle(tableStyle, tableStyleDefinitions);
            return style == null
                ? LegacyDocWritableParagraphFormatting.Plain
                : ReadSupportedStyleParagraphFormatting(style.StyleParagraphProperties);
        }

        private static LegacyDocWritableFormatting ReadSupportedTableStyleRunFormatting(TableStyle? tableStyle, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            string? styleId = tableStyle?.Val?.Value;
            if (IsNoOpTableStyle(styleId) || string.Equals(styleId, "TableGrid", StringComparison.OrdinalIgnoreCase)) {
                return LegacyDocWritableFormatting.Plain;
            }

            Style? style = ResolveSupportedTableStyle(tableStyle, tableStyleDefinitions);
            return style == null
                ? LegacyDocWritableFormatting.Plain
                : ReadSupportedRunFormatting(style.StyleRunProperties);
        }

        private static Style? ResolveSupportedTableStyle(TableStyle? tableStyle, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            string? styleId = tableStyle?.Val?.Value;
            if (IsNoOpTableStyle(styleId)) {
                return null;
            }

            if (IsTableGridStyle(styleId)) {
                return WordTableStyles.GetStyleDefinition(WordTableStyle.TableGrid);
            }

            if (string.IsNullOrWhiteSpace(styleId) || !tableStyleDefinitions.TryGetValue(styleId!, out Style? style)) {
                throw new NotSupportedException($"Native DOC saving supports simple tables only when table style '{styleId}' can be resolved to supported table-level formatting.");
            }

            ThrowIfUnsupportedTableStyle(styleId!, style, tableStyleDefinitions);
            return style;
        }

        private static bool IsNoOpTableStyle(string? styleId) {
            if (string.IsNullOrWhiteSpace(styleId)) {
                return true;
            }

            return string.Equals(styleId, "TableNormal", StringComparison.OrdinalIgnoreCase)
                || string.Equals(styleId, "NormalTable", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsTableGridStyle(string? styleId) =>
            string.Equals(styleId, "TableGrid", StringComparison.OrdinalIgnoreCase);

        private static void ThrowIfUnsupportedTableStyle(string styleId, Style style, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            if (style.Type?.Value != StyleValues.Table) {
                throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' only when it is a table style.");
            }

            foreach (OpenXmlElement child in style.ChildElements) {
                switch (child) {
                    case StyleName:
                    case UIPriority:
                    case Rsid:
                    case SemiHidden:
                    case UnhideWhenUsed:
                    case PrimaryStyle:
                    case Locked:
                    case StylePaneFormatFilter:
                        break;
                    case BasedOn basedOn:
                        ThrowIfUnsupportedTableStyleBase(styleId, basedOn, tableStyleDefinitions, new HashSet<string>(StringComparer.OrdinalIgnoreCase) { styleId });
                        break;
                    case StyleTableProperties styleTableProperties:
                        ThrowIfUnsupportedStyleTableProperties(styleId, styleTableProperties);
                        break;
                    case TableStyleProperties tableStyleProperties:
                        ThrowIfUnsupportedTableStyleConditionalProperties(styleId, tableStyleProperties);
                        break;
                    case StyleParagraphProperties styleParagraphProperties:
                        ThrowIfUnsupportedTableStyleParagraphProperties(styleId, styleParagraphProperties);
                        break;
                    case StyleRunProperties styleRunProperties:
                        ThrowIfUnsupportedTableStyleRunProperties(styleId, styleRunProperties);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving does not support table style '{styleId}' element '{child.LocalName}'.");
                }
            }
        }

        private static void ThrowIfUnsupportedTableStyleParagraphProperties(string styleId, StyleParagraphProperties paragraphProperties) {
            try {
                _ = ReadSupportedStyleParagraphFormatting(paragraphProperties);
            } catch (NotSupportedException exception) {
                throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' paragraph formatting only with supported paragraph properties. {exception.Message}", exception);
            }
        }

        private static void ThrowIfUnsupportedTableStyleRunProperties(string styleId, StyleRunProperties runProperties) {
            try {
                _ = ReadSupportedRunFormatting(runProperties);
            } catch (NotSupportedException exception) {
                throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' run formatting only with supported run properties. {exception.Message}", exception);
            }
        }

        private static void ThrowIfUnsupportedTableStyleBase(string styleId, BasedOn basedOn, IReadOnlyDictionary<string, Style> tableStyleDefinitions, ISet<string> visitedStyleIds) {
            string? baseStyleId = basedOn.Val?.Value;
            if (IsNoOpTableStyle(baseStyleId) || IsTableGridStyle(baseStyleId)) {
                return;
            }

            if (string.IsNullOrWhiteSpace(baseStyleId)
                || !tableStyleDefinitions.TryGetValue(baseStyleId!, out Style? baseStyle)) {
                throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' only when its base style can be resolved to supported table-level formatting.");
            }

            ThrowIfUnsupportedInheritedTableStyleBase(styleId, baseStyleId!, baseStyle, tableStyleDefinitions, visitedStyleIds);
        }

        private static LegacyDocTableBorders ReadSupportedTableStyleBaseBorders(Style style, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            return ReadSupportedTableStyleBaseBorders(style, tableStyleDefinitions, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
        }

        private static LegacyDocTableCellShading ReadSupportedTableStyleBaseShading(Style style, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            return ReadSupportedTableStyleBaseShading(style, tableStyleDefinitions, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
        }

        private static LegacyDocTableBorders ReadSupportedTableStyleBaseBorders(Style style, IReadOnlyDictionary<string, Style> tableStyleDefinitions, ISet<string> visitedStyleIds) {
            string? baseStyleId = style.GetFirstChild<BasedOn>()?.Val?.Value;
            if (IsNoOpTableStyle(baseStyleId)) {
                return default;
            }

            if (IsTableGridStyle(baseStyleId)) {
                return ReadSupportedTableGridBorders();
            }

            if (string.IsNullOrWhiteSpace(baseStyleId)
                || !tableStyleDefinitions.TryGetValue(baseStyleId!, out Style? baseStyle)) {
                return default;
            }

            string currentStyleId = style.StyleId?.Value ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(currentStyleId) && !visitedStyleIds.Add(currentStyleId)) {
                throw new NotSupportedException($"Native DOC saving cannot write table style '{currentStyleId}' because its basedOn chain contains a cycle.");
            }

            ThrowIfUnsupportedInheritedTableStyleBase(currentStyleId, baseStyleId!, baseStyle, tableStyleDefinitions, visitedStyleIds);
            LegacyDocTableBorders inheritedBorders = ReadSupportedTableStyleBaseBorders(baseStyle, tableStyleDefinitions, visitedStyleIds);
            LegacyDocTableBorders baseBorders = ReadSupportedTableStyleOwnBorders(baseStyle);
            return MergeSupportedTableBorders(baseBorders, inheritedBorders);
        }

        private static LegacyDocTableCellShading ReadSupportedTableStyleBaseShading(Style style, IReadOnlyDictionary<string, Style> tableStyleDefinitions, ISet<string> visitedStyleIds) {
            string? baseStyleId = style.GetFirstChild<BasedOn>()?.Val?.Value;
            if (IsNoOpTableStyle(baseStyleId)) {
                return default;
            }

            if (IsTableGridStyle(baseStyleId)) {
                return ReadSupportedTableGridShading();
            }

            if (string.IsNullOrWhiteSpace(baseStyleId)
                || !tableStyleDefinitions.TryGetValue(baseStyleId!, out Style? baseStyle)) {
                return default;
            }

            string currentStyleId = style.StyleId?.Value ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(currentStyleId) && !visitedStyleIds.Add(currentStyleId)) {
                throw new NotSupportedException($"Native DOC saving cannot write table style '{currentStyleId}' because its basedOn chain contains a cycle.");
            }

            ThrowIfUnsupportedInheritedTableStyleBase(currentStyleId, baseStyleId!, baseStyle, tableStyleDefinitions, visitedStyleIds);
            LegacyDocTableCellShading inheritedShading = ReadSupportedTableStyleBaseShading(baseStyle, tableStyleDefinitions, visitedStyleIds);
            LegacyDocTableCellShading baseShading = ReadSupportedTableStyleOwnShading(baseStyle);
            return baseShading.HasAny ? baseShading : inheritedShading;
        }

        private static void ThrowIfUnsupportedInheritedTableStyleBase(string styleId, string baseStyleId, Style baseStyle, IReadOnlyDictionary<string, Style> tableStyleDefinitions, ISet<string> visitedStyleIds) {
            if (baseStyle.Type?.Value != StyleValues.Table) {
                throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' base style '{baseStyleId}' only when it is a table style.");
            }

            if (!visitedStyleIds.Add(baseStyleId)) {
                throw new NotSupportedException($"Native DOC saving cannot write table style '{styleId}' because its basedOn chain contains a cycle.");
            }

            foreach (OpenXmlElement child in baseStyle.ChildElements) {
                switch (child) {
                    case StyleName:
                    case UIPriority:
                    case Rsid:
                    case SemiHidden:
                    case UnhideWhenUsed:
                    case PrimaryStyle:
                    case Locked:
                    case StylePaneFormatFilter:
                        break;
                    case BasedOn basedOn:
                        ThrowIfUnsupportedTableStyleBase(baseStyleId, basedOn, tableStyleDefinitions, visitedStyleIds);
                        break;
                    case StyleTableProperties styleTableProperties:
                        ThrowIfUnsupportedInheritedStyleTableProperties(styleId, baseStyleId, styleTableProperties);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' base style '{baseStyleId}' only when inherited custom formatting is supported table borders and shading. Unsupported base style element: {child.LocalName}.");
                }
            }

            visitedStyleIds.Remove(baseStyleId);
        }

        private static void ThrowIfUnsupportedInheritedStyleTableProperties(string styleId, string baseStyleId, StyleTableProperties styleTableProperties) {
            foreach (OpenXmlElement child in styleTableProperties.ChildElements) {
                switch (child) {
                    case TableBorders tableBorders:
                        ReadSupportedTableBorders(tableBorders);
                        break;
                    case Shading shading:
                        ReadSupportedTableCellShading(shading, "inherited table style shading");
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' base style '{baseStyleId}' only when inherited custom table-level formatting is supported borders and shading. Unsupported inherited table style property: {child.LocalName}.");
                }
            }
        }

        private static LegacyDocTableBorders ReadSupportedTableStyleOwnBorders(Style style) {
            TableBorders? borders = style.GetFirstChild<StyleTableProperties>()?.GetFirstChild<TableBorders>();
            return borders == null ? default : ReadSupportedTableBorders(borders);
        }

        private static LegacyDocTableCellShading ReadSupportedTableStyleOwnShading(Style style) {
            Shading? shading = style.GetFirstChild<StyleTableProperties>()?.GetFirstChild<Shading>();
            return shading == null ? default : ReadSupportedTableCellShading(shading, "table style shading");
        }

        private static LegacyDocTableBorders ReadSupportedTableGridBorders() {
            Style styleDefinition = WordTableStyles.GetStyleDefinition(WordTableStyle.TableGrid);
            TableBorders? borders = styleDefinition.GetFirstChild<StyleTableProperties>()?.GetFirstChild<TableBorders>();
            return borders == null ? default : ReadSupportedTableBorders(borders);
        }

        private static LegacyDocTableCellShading ReadSupportedTableGridShading() {
            Style styleDefinition = WordTableStyles.GetStyleDefinition(WordTableStyle.TableGrid);
            Shading? shading = styleDefinition.GetFirstChild<StyleTableProperties>()?.GetFirstChild<Shading>();
            return shading == null ? default : ReadSupportedTableCellShading(shading, "table style shading");
        }

        private static LegacyDocTableBorders MergeSupportedTableBorders(LegacyDocTableBorders ownBorders, LegacyDocTableBorders inheritedBorders) {
            if (!inheritedBorders.HasAny || !ownBorders.HasAny) {
                return ownBorders.HasAny ? ownBorders : inheritedBorders;
            }

            return new LegacyDocTableBorders(
                ownBorders.Top.HasAny ? ownBorders.Top : inheritedBorders.Top,
                ownBorders.Left.HasAny ? ownBorders.Left : inheritedBorders.Left,
                ownBorders.Bottom.HasAny ? ownBorders.Bottom : inheritedBorders.Bottom,
                ownBorders.Right.HasAny ? ownBorders.Right : inheritedBorders.Right,
                ownBorders.InsideHorizontal.HasAny ? ownBorders.InsideHorizontal : inheritedBorders.InsideHorizontal,
                ownBorders.InsideVertical.HasAny ? ownBorders.InsideVertical : inheritedBorders.InsideVertical);
        }

        private static void ThrowIfUnsupportedStyleTableProperties(string styleId, StyleTableProperties styleTableProperties) {
            foreach (OpenXmlElement child in styleTableProperties.ChildElements) {
                switch (child) {
                    case TableBorders tableBorders:
                        ReadSupportedTableBorders(tableBorders);
                        break;
                    case Shading shading:
                        ReadSupportedTableCellShading(shading, "table style shading");
                        break;
                    case TableCellMarginDefault tableCellMarginDefault:
                        ReadSupportedTableDefaultCellMargins(tableCellMarginDefault);
                        break;
                    case TableCellSpacing tableCellSpacing:
                        ReadSupportedTableDefaultCellSpacing(tableCellSpacing);
                        break;
                    case TableJustification tableJustification:
                        ReadSupportedTableAlignment(tableJustification);
                        break;
                    case TableIndentation tableIndentation:
                        ReadSupportedTableIndentation(tableIndentation);
                        break;
                    case TableWidth tableWidth:
                        ReadSupportedTablePreferredWidth(tableWidth);
                        break;
                    case TableLayout tableLayout:
                        ReadSupportedTableAutofit(tableLayout);
                        break;
                    case TableStyleRowBandSize rowBandSize:
                        ReadSupportedTableStyleBandSize(rowBandSize.Val, "row");
                        break;
                    case TableStyleColumnBandSize columnBandSize:
                        ReadSupportedTableStyleBandSize(columnBandSize.Val, "column");
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' only with supported table-level layout, borders, shading, default cell margins, and default cell spacing. Unsupported table style property: {child.LocalName}.");
                }
            }
        }

        private static void ThrowIfUnsupportedTableStyleConditionalProperties(string styleId, TableStyleProperties tableStyleProperties) {
            if (tableStyleProperties.Type?.Value == null) {
                throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' conditional formatting only when the conditional type is specified.");
            }

            foreach (OpenXmlElement child in tableStyleProperties.ChildElements) {
                switch (child) {
                    case TableStyleConditionalFormattingTableCellProperties cellProperties:
                        ThrowIfUnsupportedTableStyleConditionalCellProperties(styleId, cellProperties);
                        break;
                    case TableStyleConditionalFormattingTableProperties tableProperties:
                        ThrowIfUnsupportedTableStyleConditionalTableProperties(styleId, tableProperties);
                        break;
                    case StyleParagraphProperties styleParagraphProperties:
                        ThrowIfUnsupportedTableStyleConditionalParagraphProperties(styleId, styleParagraphProperties);
                        break;
                    case StyleRunProperties styleRunProperties:
                        ThrowIfUnsupportedTableStyleConditionalRunProperties(styleId, styleRunProperties);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' conditional formatting only with supported table, cell, paragraph, and run effects. Unsupported conditional style element: {child.LocalName}.");
                }
            }
        }

        private static void ThrowIfUnsupportedTableStyleConditionalParagraphProperties(string styleId, StyleParagraphProperties paragraphProperties) {
            try {
                _ = ReadSupportedStyleParagraphFormatting(paragraphProperties);
            } catch (NotSupportedException exception) {
                throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' conditional paragraph formatting only with supported paragraph properties. {exception.Message}", exception);
            }
        }

        private static void ThrowIfUnsupportedTableStyleConditionalRunProperties(string styleId, StyleRunProperties runProperties) {
            try {
                _ = ReadSupportedRunFormatting(runProperties);
            } catch (NotSupportedException exception) {
                throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' conditional run formatting only with supported run properties. {exception.Message}", exception);
            }
        }

        private static void ThrowIfUnsupportedTableStyleConditionalCellProperties(string styleId, TableStyleConditionalFormattingTableCellProperties cellProperties) {
            foreach (OpenXmlElement child in cellProperties.ChildElements) {
                switch (child) {
                    case TableCellVerticalAlignment verticalAlignment:
                        ReadSupportedTableCellVerticalAlignment(verticalAlignment);
                        break;
                    case TextDirection textDirection:
                        ReadSupportedTableCellTextDirection(textDirection);
                        break;
                    case TableCellFitText fitText:
                        ReadSupportedTableCellFitText(fitText);
                        break;
                    case NoWrap noWrap:
                        ReadSupportedTableCellNoWrap(noWrap);
                        break;
                    case HideMark hideMark:
                        ReadSupportedTableCellHideMark(hideMark);
                        break;
                    case TableCellMargin margins:
                        ReadSupportedTableCellMargins(margins);
                        break;
                    case Shading shading:
                        ReadSupportedTableCellShading(shading, "conditional table style shading");
                        break;
                    case TableCellBorders borders:
                        ThrowIfUnsupportedTableStyleConditionalCellBorders(styleId, borders);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' conditional cell formatting only with supported layout, borders, and shading. Unsupported conditional cell property: {child.LocalName}.");
                }
            }
        }

        private static void ThrowIfUnsupportedTableStyleConditionalCellBorders(string styleId, TableCellBorders borders) {
            foreach (OpenXmlElement child in borders.ChildElements) {
                switch (child) {
                    case TopBorder:
                    case LeftBorder:
                    case BottomBorder:
                    case RightBorder:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' conditional borders only on top, left, bottom, and right cell edges. Unsupported conditional border: {child.LocalName}.");
                }
            }

            ReadSupportedTableCellBorder(borders.TopBorder);
            ReadSupportedTableCellBorder(borders.LeftBorder);
            ReadSupportedTableCellBorder(borders.BottomBorder);
            ReadSupportedTableCellBorder(borders.RightBorder);
        }

        private static void ThrowIfUnsupportedTableStyleConditionalTableProperties(string styleId, TableStyleConditionalFormattingTableProperties tableProperties) {
            foreach (OpenXmlElement child in tableProperties.ChildElements) {
                switch (child) {
                    case TableBorders tableBorders:
                        ReadSupportedTableBorders(tableBorders);
                        break;
                    case Shading shading:
                        ReadSupportedTableCellShading(shading, "conditional table style shading");
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' conditional table formatting only with supported borders and shading. Unsupported conditional table property: {child.LocalName}.");
                }
            }
        }
    }
}
