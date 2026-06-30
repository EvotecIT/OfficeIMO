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

            if (string.Equals(styleId, "TableGrid", StringComparison.OrdinalIgnoreCase)) {
                Style styleDefinition = WordTableStyles.GetStyleDefinition(WordTableStyle.TableGrid);
                TableBorders? borders = styleDefinition.GetFirstChild<StyleTableProperties>()?.GetFirstChild<TableBorders>();
                return borders == null ? default : ReadSupportedTableBorders(borders);
            }

            if (string.IsNullOrWhiteSpace(styleId) || !tableStyleDefinitions.TryGetValue(styleId!, out Style? style)) {
                throw new NotSupportedException($"Native DOC saving supports simple tables only when table style '{styleId}' can be resolved to supported table-level formatting.");
            }

            ThrowIfUnsupportedTableStyle(styleId!, style);
            TableBorders? customBorders = style.GetFirstChild<StyleTableProperties>()?.GetFirstChild<TableBorders>();
            return customBorders == null ? default : ReadSupportedTableBorders(customBorders);
        }

        private static LegacyDocTableCellShading ReadSupportedTableStyleShading(TableStyle? tableStyle, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            string? styleId = tableStyle?.Val?.Value;
            if (IsNoOpTableStyle(styleId)) {
                return default;
            }

            if (string.Equals(styleId, "TableGrid", StringComparison.OrdinalIgnoreCase)) {
                Style styleDefinition = WordTableStyles.GetStyleDefinition(WordTableStyle.TableGrid);
                Shading? shading = styleDefinition.GetFirstChild<StyleTableProperties>()?.GetFirstChild<Shading>();
                return shading == null ? default : ReadSupportedTableCellShading(shading, "table style shading");
            }

            if (string.IsNullOrWhiteSpace(styleId) || !tableStyleDefinitions.TryGetValue(styleId!, out Style? style)) {
                throw new NotSupportedException($"Native DOC saving supports simple tables only when table style '{styleId}' can be resolved to supported table-level formatting.");
            }

            ThrowIfUnsupportedTableStyle(styleId!, style);
            Shading? customShading = style.GetFirstChild<StyleTableProperties>()?.GetFirstChild<Shading>();
            return customShading == null ? default : ReadSupportedTableCellShading(customShading, "table style shading");
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

        private static Style? ResolveSupportedTableStyle(TableStyle? tableStyle, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            string? styleId = tableStyle?.Val?.Value;
            if (IsNoOpTableStyle(styleId)) {
                return null;
            }

            if (string.Equals(styleId, "TableGrid", StringComparison.OrdinalIgnoreCase)) {
                return WordTableStyles.GetStyleDefinition(WordTableStyle.TableGrid);
            }

            if (string.IsNullOrWhiteSpace(styleId) || !tableStyleDefinitions.TryGetValue(styleId!, out Style? style)) {
                throw new NotSupportedException($"Native DOC saving supports simple tables only when table style '{styleId}' can be resolved to supported table-level formatting.");
            }

            ThrowIfUnsupportedTableStyle(styleId!, style);
            return style;
        }

        private static bool IsNoOpTableStyle(string? styleId) {
            if (string.IsNullOrWhiteSpace(styleId)) {
                return true;
            }

            return string.Equals(styleId, "TableNormal", StringComparison.OrdinalIgnoreCase)
                || string.Equals(styleId, "NormalTable", StringComparison.OrdinalIgnoreCase);
        }

        private static void ThrowIfUnsupportedTableStyle(string styleId, Style style) {
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
                        ThrowIfUnsupportedTableStyleBase(styleId, basedOn);
                        break;
                    case StyleTableProperties styleTableProperties:
                        ThrowIfUnsupportedStyleTableProperties(styleId, styleTableProperties);
                        break;
                    case StyleParagraphProperties:
                    case StyleRunProperties:
                    case TableStyleProperties:
                        throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' only when it contains table-level border formatting without conditional, paragraph, or run style effects.");
                    default:
                        throw new NotSupportedException($"Native DOC saving does not support table style '{styleId}' element '{child.LocalName}'.");
                }
            }
        }

        private static void ThrowIfUnsupportedTableStyleBase(string styleId, BasedOn basedOn) {
            string? baseStyleId = basedOn.Val?.Value;
            if (IsNoOpTableStyle(baseStyleId)) {
                return;
            }

            throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' only when it is based on TableNormal.");
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
                    default:
                        throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' only with table-level borders, shading, default cell margins, and default cell spacing. Unsupported table style property: {child.LocalName}.");
                }
            }
        }
    }
}
