using OfficeIMO.Drawing;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        private static bool TryReadImageExportTitle(OfficeXmlNode chart, ExcelPartRef workbook,
            out OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartTextReader.NativeText text) {
            text = default;
            OfficeXmlNode title = chart.Child(C, "title");
            if (title.IsNone) return false;
            try {
                text = ImageExportReadNativeText(chart, title,
                    workbook.FirstPartOfKind(ExcelPartKind.Theme).Root.Child(A, "themeElements").Child(A, "clrScheme"));
                return true;
            } catch (NotSupportedException) {
                return false;
            }
        }

        private static bool TryGetImageExportTitleColor(OfficeXmlNode chart, ExcelPartRef workbookPart, out OfficeColor color) {
            bool resolved = TryReadImageExportTitle(chart, workbookPart, out var text);
            color = text.Color.GetValueOrDefault();
            return resolved && text.Color.HasValue;
        }

        private static bool TryGetImageExportTitleFontFamily(OfficeXmlNode chart, ExcelPartRef workbookPart, out string? fontFamily) {
            bool resolved = TryReadImageExportTitle(chart, workbookPart, out var text);
            fontFamily = text.Family;
            return resolved && fontFamily != null;
        }

        private static bool TryGetImageExportTitleFontSize(OfficeXmlNode chart, ExcelPartRef workbookPart, out double fontSize) {
            bool resolved = TryReadImageExportTitle(chart, workbookPart, out var text);
            fontSize = text.Size.GetValueOrDefault();
            return resolved && text.Size.HasValue;
        }

        private static bool TryGetImageExportTitleFontStyle(OfficeXmlNode chart, ExcelPartRef workbookPart, out OfficeFontStyle fontStyle) {
            bool resolved = TryReadImageExportTitle(chart, workbookPart, out var text);
            fontStyle = text.Style.GetValueOrDefault();
            return resolved && text.Style.HasValue;
        }

        private static bool TryGetImageExportLegendTextColor(OfficeXmlNode chart, ExcelPartRef workbookPart, out OfficeColor color) =>
            TryGetFirstImageExportTextColor(GetImageExportLegendTextOwners(chart), workbookPart, out color);

        private static bool TryGetImageExportDataLabelTextColor(OfficeXmlNode plotArea, ExcelPartRef workbookPart, out OfficeColor color) =>
            TryGetFirstImageExportTextColor(GetImageExportDataLabelTextOwners(plotArea), workbookPart, out color);

        private static bool TryGetImageExportDataLabelFillColor(OfficeXmlNode plotArea, ExcelPartRef workbookPart, out OfficeColor color) =>
            TryGetFirstImageExportShapeFill(GetImageExportDataLabelTextOwners(plotArea), workbookPart, out color);

        private static bool TryGetImageExportDataLabelLineColor(OfficeXmlNode plotArea, ExcelPartRef workbookPart, out OfficeColor color) =>
            TryGetFirstImageExportShapeLine(GetImageExportDataLabelTextOwners(plotArea), workbookPart, out color);

        private static bool TryGetImageExportDataLabelLineWidth(OfficeXmlNode plotArea, out double width) =>
            TryGetFirstImageExportShapeLineWidth(GetImageExportDataLabelTextOwners(plotArea), out width);

        private static bool TryGetImageExportDataLabelLineDashStyle(OfficeXmlNode plotArea, out OfficeStrokeDashStyle dashStyle) =>
            TryGetFirstImageExportShapeLineDashStyle(GetImageExportDataLabelTextOwners(plotArea), out dashStyle);

        private static bool TryGetImageExportAxisTextColor(OfficeXmlNode plotArea, ExcelPartRef workbookPart, out OfficeColor color) =>
            TryGetFirstImageExportAxisLabelTextColor(plotArea, workbookPart, out color);

        private static bool TryGetImageExportAxisTitleTextColor(OfficeXmlNode plotArea, ExcelPartRef workbookPart, out OfficeColor color) =>
            TryGetFirstImageExportTextColor(GetImageExportAxisTitleTextOwners(plotArea), workbookPart, out color);

        private static bool TryGetImageExportLegendFontSize(OfficeXmlNode chart, out double fontSize) =>
            TryGetFirstImageExportTextFontSize(GetImageExportLegendTextOwners(chart), out fontSize);

        private static bool TryGetImageExportLegendFontFamily(OfficeXmlNode chart, out string? fontFamily) =>
            TryGetFirstImageExportTextFontFamily(GetImageExportLegendTextOwners(chart), out fontFamily);

        private static bool TryGetImageExportLegendFontStyle(OfficeXmlNode chart, out OfficeFontStyle fontStyle) =>
            TryGetFirstImageExportTextFontStyle(GetImageExportLegendTextOwners(chart), out fontStyle);

        private static bool TryGetImageExportDataLabelFontSize(OfficeXmlNode plotArea, out double fontSize) =>
            TryGetFirstImageExportTextFontSize(GetImageExportDataLabelTextOwners(plotArea), out fontSize);

        private static bool TryGetImageExportDataLabelFontFamily(OfficeXmlNode plotArea, out string? fontFamily) =>
            TryGetFirstImageExportTextFontFamily(GetImageExportDataLabelTextOwners(plotArea), out fontFamily);

        private static bool TryGetImageExportDataLabelFontStyle(OfficeXmlNode plotArea, out OfficeFontStyle fontStyle) =>
            TryGetFirstImageExportTextFontStyle(GetImageExportDataLabelTextOwners(plotArea), out fontStyle);

        private static bool TryGetImageExportAxisLabelFontSize(OfficeXmlNode plotArea, out double fontSize) =>
            TryGetFirstImageExportAxisLabelTextFontSize(plotArea, out fontSize);

        private static bool TryGetImageExportAxisTitleFontSize(OfficeXmlNode plotArea, out double fontSize) =>
            TryGetFirstImageExportTextFontSize(GetImageExportAxisTitleTextOwners(plotArea), out fontSize);

        private static bool TryGetImageExportAxisTextFontFamily(OfficeXmlNode plotArea, out string? fontFamily) =>
            TryGetFirstImageExportAxisLabelTextFontFamily(plotArea, out fontFamily);

        private static bool TryGetImageExportAxisTitleFontFamily(OfficeXmlNode plotArea, out string? fontFamily) =>
            TryGetFirstImageExportTextFontFamily(GetImageExportAxisTitleTextOwners(plotArea), out fontFamily);

        private static bool TryGetImageExportAxisTextFontStyle(OfficeXmlNode plotArea, out OfficeFontStyle fontStyle) =>
            TryGetFirstImageExportAxisLabelTextFontStyle(plotArea, out fontStyle);

        private static bool TryGetImageExportAxisTitleFontStyle(OfficeXmlNode plotArea, out OfficeFontStyle fontStyle) =>
            TryGetFirstImageExportTextFontStyle(GetImageExportAxisTitleTextOwners(plotArea), out fontStyle);

        private static bool HasUnsupportedImageExportTextStyle(OfficeXmlNode chartSpace, ExcelPartRef workbookPart) {
            OfficeXmlNode chart = chartSpace.Child(C, "chart");
            OfficeXmlNode plotArea = chart.Child(C, "plotArea");
            OfficeXmlNode chartTitle = chart.Child(C, "title");
            if (chartTitle.Exists && !TryReadImageExportTitle(chart, workbookPart, out _)) {
                return true;
            }

            IReadOnlyList<OfficeXmlNode> bodyOwners = GetImageExportBodyTextOwners(chart, plotArea).ToArray();
            foreach (OfficeXmlNode owner in bodyOwners) {
                if (HasUnsupportedImageExportTextStyle(owner, allowSolidFill: true, allowFontFamily: true, allowFontSize: true, allowFontStyle: true, workbookPart)) {
                    return true;
                }
            }

            IReadOnlyList<OfficeXmlNode> axisLabelOwners = GetImageExportAxisLabelTextOwners(plotArea).ToArray();
            foreach (OfficeXmlNode owner in axisLabelOwners) {
                if (HasUnsupportedImageExportAxisLabelTextStyle(owner, allowSolidFill: true, allowFontFamily: true, allowFontSize: true, allowFontStyle: true, workbookPart)) {
                    return true;
                }
            }

            IReadOnlyList<OfficeXmlNode> axisTitleOwners = GetImageExportAxisTitleTextOwners(plotArea).ToArray();
            foreach (OfficeXmlNode owner in axisTitleOwners) {
                if (HasUnsupportedImageExportTextStyle(owner, allowSolidFill: true, allowFontFamily: true, allowFontSize: true, allowFontStyle: true, workbookPart)) {
                    return true;
                }
            }

            IReadOnlyList<OfficeXmlNode> legendOwners = GetImageExportLegendTextOwners(chart).ToArray();
            IReadOnlyList<OfficeXmlNode> dataLabelOwners = GetImageExportDataLabelTextOwners(plotArea).ToArray();
            return HasConflictingImageExportTextColors(legendOwners, workbookPart) ||
                HasConflictingImageExportTextColors(dataLabelOwners, workbookPart) ||
                HasConflictingImageExportAxisLabelTextColors(axisLabelOwners, workbookPart) ||
                HasConflictingImageExportTextColors(axisTitleOwners, workbookPart) ||
                HasConflictingImageExportTextFontFamilies(legendOwners) ||
                HasConflictingImageExportTextFontFamilies(dataLabelOwners) ||
                HasConflictingImageExportAxisLabelTextFontFamilies(axisLabelOwners) ||
                HasConflictingImageExportTextFontFamilies(axisTitleOwners) ||
                HasConflictingImageExportTextFontSizes(dataLabelOwners) ||
                HasConflictingImageExportAxisLabelTextFontSizes(axisLabelOwners) ||
                HasConflictingImageExportTextFontSizes(axisTitleOwners) ||
                HasConflictingImageExportTextFontStyles(legendOwners) ||
                HasConflictingImageExportTextFontStyles(dataLabelOwners) ||
                HasConflictingImageExportAxisLabelTextFontStyles(axisLabelOwners) ||
                HasConflictingImageExportTextFontStyles(axisTitleOwners);
        }

        private static IEnumerable<OfficeXmlNode> GetImageExportLegendTextOwners(OfficeXmlNode chart) {
            OfficeXmlNode legend = chart.Child(C, "legend");
            if (legend.Exists) {
                yield return legend;
            }
        }

        private static IEnumerable<OfficeXmlNode> GetImageExportDataLabelTextOwners(OfficeXmlNode plotArea) {
            if (plotArea.IsNone) {
                yield break;
            }

            foreach (OfficeXmlNode labels in plotArea.Descendants(C, "dLbls")) {
                if (!IsImageExportSchemaPlaced(labels)) continue;
                yield return labels;
            }
        }

        private static IEnumerable<OfficeXmlNode> GetImageExportBodyTextOwners(OfficeXmlNode chart, OfficeXmlNode plotArea) {
            foreach (OfficeXmlNode legend in GetImageExportLegendTextOwners(chart)) {
                yield return legend;
            }

            foreach (OfficeXmlNode labels in GetImageExportDataLabelTextOwners(plotArea)) {
                yield return labels;
            }
        }

        private static bool TryGetFirstImageExportShapeFill(IEnumerable<OfficeXmlNode> owners, ExcelPartRef workbookPart, out OfficeColor color) {
            foreach (OfficeXmlNode owner in owners) {
                if (TryGetSolidFill(owner.Child(C, "spPr"), workbookPart, out color)) {
                    return true;
                }
            }

            color = default;
            return false;
        }

        private static bool TryGetFirstImageExportShapeLine(IEnumerable<OfficeXmlNode> owners, ExcelPartRef workbookPart, out OfficeColor color) {
            foreach (OfficeXmlNode owner in owners) {
                if (TryGetSolidLine(owner.Child(C, "spPr"), workbookPart, out color)) {
                    return true;
                }
            }

            color = default;
            return false;
        }

        private static bool TryGetFirstImageExportShapeLineWidth(IEnumerable<OfficeXmlNode> owners, out double width) {
            foreach (OfficeXmlNode owner in owners) {
                if (TryGetImageExportLineWidth(owner.Child(C, "spPr"), out width)) {
                    return true;
                }
            }

            width = default;
            return false;
        }

        private static bool TryGetFirstImageExportShapeLineDashStyle(IEnumerable<OfficeXmlNode> owners, out OfficeStrokeDashStyle dashStyle) {
            foreach (OfficeXmlNode owner in owners) {
                if (TryGetLineDashStyle(owner.Child(C, "spPr"), out dashStyle)) {
                    return true;
                }
            }

            dashStyle = default;
            return false;
        }

        private static IEnumerable<OfficeXmlNode> GetImageExportAxisTextOwners(OfficeXmlNode plotArea) {
            foreach (OfficeXmlNode axis in GetImageExportAxisLabelTextOwners(plotArea)) {
                yield return axis;
            }

            foreach (OfficeXmlNode title in GetImageExportAxisTitleTextOwners(plotArea)) {
                yield return title;
            }
        }

        private static IEnumerable<OfficeXmlNode> GetImageExportAxisLabelTextOwners(OfficeXmlNode plotArea) {
            if (plotArea.IsNone) {
                yield break;
            }

            foreach (OfficeXmlNode axis in GetImageExportAxes(plotArea)) {
                yield return axis;
            }
        }

        private static IEnumerable<OfficeXmlNode> GetImageExportAxisTitleTextOwners(OfficeXmlNode plotArea) {
            if (plotArea.IsNone) {
                yield break;
            }

            foreach (OfficeXmlNode axis in GetImageExportAxes(plotArea)) {
                OfficeXmlNode title = axis.Child(C, "title");
                if (title.Exists) {
                    yield return title;
                }
            }
        }

        private static bool TryGetFirstImageExportTextColor(IEnumerable<OfficeXmlNode> owners, ExcelPartRef workbookPart, out OfficeColor color) {
            foreach (OfficeXmlNode owner in owners) {
                foreach (OfficeXmlNode properties in GetImageExportTextRunProperties(owner)) {
                    if (TryGetSolidFill(properties, workbookPart, out color)) {
                        return true;
                    }
                }
            }

            color = default;
            return false;
        }

        private static bool TryGetFirstImageExportAxisLabelTextColor(OfficeXmlNode plotArea, ExcelPartRef workbookPart, out OfficeColor color) {
            if (plotArea.Exists) {
                foreach (OfficeXmlNode axis in GetImageExportAxes(plotArea)) {
                    foreach (OfficeXmlNode properties in GetImageExportAxisLabelTextRunProperties(axis)) {
                        if (TryGetSolidFill(properties, workbookPart, out color)) {
                            return true;
                        }
                    }
                }
            }

            color = default;
            return false;
        }

        private static bool TryGetFirstImageExportTextFontFamily(IEnumerable<OfficeXmlNode> owners, out string? fontFamily) {
            foreach (OfficeXmlNode owner in owners) {
                foreach (OfficeXmlNode properties in GetImageExportTextRunProperties(owner)) {
                    if (TryGetImageExportTextFontFamily(properties, out fontFamily)) {
                        return true;
                    }
                }
            }

            fontFamily = default;
            return false;
        }

        private static bool TryGetFirstImageExportTextFontSize(IEnumerable<OfficeXmlNode> owners, out double fontSize) {
            foreach (OfficeXmlNode owner in owners) {
                foreach (OfficeXmlNode properties in GetImageExportTextRunProperties(owner)) {
                    if (TryGetImageExportTextFontSize(properties, out fontSize)) {
                        return true;
                    }
                }
            }

            fontSize = default;
            return false;
        }

        private static bool TryGetFirstImageExportAxisLabelTextFontSize(OfficeXmlNode plotArea, out double fontSize) {
            if (plotArea.Exists) {
                foreach (OfficeXmlNode axis in GetImageExportAxes(plotArea)) {
                    foreach (OfficeXmlNode properties in GetImageExportAxisLabelTextRunProperties(axis)) {
                        if (TryGetImageExportTextFontSize(properties, out fontSize)) {
                            return true;
                        }
                    }
                }
            }

            fontSize = default;
            return false;
        }

        private static bool TryGetFirstImageExportAxisLabelTextFontFamily(OfficeXmlNode plotArea, out string? fontFamily) {
            if (plotArea.Exists) {
                foreach (OfficeXmlNode axis in GetImageExportAxes(plotArea)) {
                    foreach (OfficeXmlNode properties in GetImageExportAxisLabelTextRunProperties(axis)) {
                        if (TryGetImageExportTextFontFamily(properties, out fontFamily)) {
                            return true;
                        }
                    }
                }
            }

            fontFamily = default;
            return false;
        }

        private static bool TryGetFirstImageExportTextFontStyle(IEnumerable<OfficeXmlNode> owners, out OfficeFontStyle fontStyle) {
            foreach (OfficeXmlNode owner in owners) {
                foreach (OfficeXmlNode properties in GetImageExportTextRunProperties(owner)) {
                    if (TryGetImageExportTextFontStyle(properties, out fontStyle)) {
                        return true;
                    }
                }
            }

            fontStyle = default;
            return false;
        }

        private static bool TryGetFirstImageExportAxisLabelTextFontStyle(OfficeXmlNode plotArea, out OfficeFontStyle fontStyle) {
            if (plotArea.Exists) {
                foreach (OfficeXmlNode axis in GetImageExportAxes(plotArea)) {
                    foreach (OfficeXmlNode properties in GetImageExportAxisLabelTextRunProperties(axis)) {
                        if (TryGetImageExportTextFontStyle(properties, out fontStyle)) {
                            return true;
                        }
                    }
                }
            }

            fontStyle = default;
            return false;
        }

        private static bool HasConflictingImageExportTextColors(IEnumerable<OfficeXmlNode> owners, ExcelPartRef workbookPart) {
            string? first = null;
            foreach (OfficeXmlNode owner in owners) {
                foreach (OfficeXmlNode properties in GetImageExportTextRunProperties(owner)) {
                    if (!TryGetSolidFill(properties, workbookPart, out OfficeColor color)) {
                        continue;
                    }

                    string current = color.ToRgbHex();
                    if (first == null) {
                        first = current;
                    } else if (!string.Equals(first, current, System.StringComparison.OrdinalIgnoreCase)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasConflictingImageExportAxisLabelTextColors(IEnumerable<OfficeXmlNode> owners, ExcelPartRef workbookPart) {
            string? first = null;
            foreach (OfficeXmlNode owner in owners) {
                foreach (OfficeXmlNode properties in GetImageExportAxisLabelTextRunProperties(owner)) {
                    if (!TryGetSolidFill(properties, workbookPart, out OfficeColor color)) {
                        continue;
                    }

                    string current = color.ToRgbHex();
                    if (first == null) {
                        first = current;
                    } else if (!string.Equals(first, current, System.StringComparison.OrdinalIgnoreCase)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasConflictingImageExportTextFontFamilies(IEnumerable<OfficeXmlNode> owners) {
            string? first = null;
            foreach (OfficeXmlNode owner in owners) {
                foreach (OfficeXmlNode properties in GetImageExportTextRunProperties(owner)) {
                    if (!TryGetImageExportTextFontFamily(properties, out string? current) || string.IsNullOrWhiteSpace(current)) {
                        continue;
                    }

                    if (first == null) {
                        first = current;
                    } else if (!string.Equals(first, current, StringComparison.OrdinalIgnoreCase)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasConflictingImageExportTextFontSizes(IEnumerable<OfficeXmlNode> owners) {
            double? first = null;
            foreach (OfficeXmlNode owner in owners) {
                foreach (OfficeXmlNode properties in GetImageExportTextRunProperties(owner)) {
                    if (!TryGetImageExportTextFontSize(properties, out double current)) {
                        continue;
                    }

                    if (first == null) {
                        first = current;
                    } else if (Math.Abs(first.Value - current) > 0.01D) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasConflictingImageExportAxisLabelTextFontFamilies(IEnumerable<OfficeXmlNode> owners) {
            string? first = null;
            foreach (OfficeXmlNode owner in owners) {
                foreach (OfficeXmlNode properties in GetImageExportAxisLabelTextRunProperties(owner)) {
                    if (!TryGetImageExportTextFontFamily(properties, out string? current) || string.IsNullOrWhiteSpace(current)) {
                        continue;
                    }

                    if (first == null) {
                        first = current;
                    } else if (!string.Equals(first, current, StringComparison.OrdinalIgnoreCase)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasConflictingImageExportAxisLabelTextFontSizes(IEnumerable<OfficeXmlNode> owners) {
            double? first = null;
            foreach (OfficeXmlNode owner in owners) {
                foreach (OfficeXmlNode properties in GetImageExportAxisLabelTextRunProperties(owner)) {
                    if (!TryGetImageExportTextFontSize(properties, out double current)) {
                        continue;
                    }

                    if (first == null) {
                        first = current;
                    } else if (Math.Abs(first.Value - current) > 0.01D) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasConflictingImageExportAxisLabelTextFontStyles(IEnumerable<OfficeXmlNode> owners) {
            OfficeFontStyle? first = null;
            foreach (OfficeXmlNode owner in owners) {
                foreach (OfficeXmlNode properties in GetImageExportAxisLabelTextRunProperties(owner)) {
                    if (!TryGetImageExportTextFontStyle(properties, out OfficeFontStyle current)) {
                        continue;
                    }

                    if (first == null) {
                        first = current;
                    } else if (first.Value != current) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasConflictingImageExportTextFontStyles(IEnumerable<OfficeXmlNode> owners) {
            OfficeFontStyle? first = null;
            foreach (OfficeXmlNode owner in owners) {
                foreach (OfficeXmlNode properties in GetImageExportTextRunProperties(owner)) {
                    if (!TryGetImageExportTextFontStyle(properties, out OfficeFontStyle current)) {
                        continue;
                    }

                    if (first == null) {
                        first = current;
                    } else if (first.Value != current) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasUnsupportedImageExportTextStyle(OfficeXmlNode owner, bool allowSolidFill, bool allowFontFamily, bool allowFontSize, bool allowFontStyle, ExcelPartRef workbookPart) {
            foreach (OfficeXmlNode properties in GetImageExportTextRunProperties(owner)) {
                if (!IsSimpleSupportedImageExportTextProperties(properties, allowSolidFill, allowFontFamily, allowFontSize, allowFontStyle, workbookPart)) {
                    return true;
                }
            }

            return false;
        }

        private static bool HasUnsupportedImageExportAxisLabelTextStyle(OfficeXmlNode owner, bool allowSolidFill, bool allowFontFamily, bool allowFontSize, bool allowFontStyle, ExcelPartRef workbookPart) {
            foreach (OfficeXmlNode properties in GetImageExportAxisLabelTextRunProperties(owner)) {
                if (!IsSimpleSupportedImageExportTextProperties(properties, allowSolidFill, allowFontFamily, allowFontSize, allowFontStyle, workbookPart)) {
                    return true;
                }
            }

            return false;
        }

        // The text run properties read here are a:rPr and a:defRPr, both character properties.
        private static bool IsSimpleSupportedImageExportTextProperties(OfficeXmlNode properties, bool allowSolidFill, bool allowFontFamily, bool allowFontSize, bool allowFontStyle, ExcelPartRef workbookPart) {
            if ((properties.Attribute("b") != null || properties.Attribute("i") != null) &&
                !allowFontStyle) {
                return false;
            }
            if (properties.Attribute("sz") != null) {
                if (!allowFontSize || !TryGetImageExportTextFontSize(properties, out _)) {
                    return false;
                }
            }

            foreach (OfficeXmlNode child in properties.Children()) {
                if (allowSolidFill && child.Is(A, "solidFill")) {
                    if (!TryGetSolidFill(properties, workbookPart, out _)) {
                        return false;
                    }

                    continue;
                }

                if (allowFontFamily && child.Is(A, "latin")) {
                    if (!TryGetImageExportTextFontFamily(properties, out _)) {
                        return false;
                    }

                    continue;
                }

                return false;
            }

            return true;
        }

        private static bool TryGetImageExportTextFontFamily(OfficeXmlNode properties, out string? fontFamily) {
            string? typeface = properties.Child(A, "latin").Attribute("typeface");
            if (!string.IsNullOrWhiteSpace(typeface)) {
                fontFamily = typeface;
                return true;
            }

            fontFamily = default;
            return false;
        }

        private static bool IsThemeFontFamilyPlaceholder(string fontFamily) =>
            fontFamily.Length > 0 && fontFamily[0] == '+';

        private static bool TryGetImageExportTextFontSize(OfficeXmlNode properties, out double fontSize) {
            if (properties.Int32Attribute("sz") is int size && size > 0) {
                fontSize = size / 100D;
                return true;
            }

            fontSize = default;
            return false;
        }

        private static bool TryGetImageExportTextFontStyle(OfficeXmlNode properties, out OfficeFontStyle fontStyle) {
            fontStyle = OfficeFontStyle.Regular;
            bool hasStyle = false;
            // A value that isn't an xsd:boolean throws, as the SDK's typed read did, so the snapshot reports the style.
            if (properties.Attribute("b") is string bold) {
                hasStyle = true;
                if (System.Xml.XmlConvert.ToBoolean(bold)) {
                    fontStyle |= OfficeFontStyle.Bold;
                }
            }
            if (properties.Attribute("i") is string italic) {
                hasStyle = true;
                if (System.Xml.XmlConvert.ToBoolean(italic)) {
                    fontStyle |= OfficeFontStyle.Italic;
                }
            }

            return hasStyle;
        }

        private static IEnumerable<OfficeXmlNode> GetImageExportTextRunProperties(OfficeXmlNode owner) =>
            owner.Descendants(A, "rPr")
                .Concat(owner.Descendants(A, "defRPr"));

        private static IEnumerable<OfficeXmlNode> GetImageExportAxisLabelTextRunProperties(OfficeXmlNode owner) =>
            owner.Descendants(A, "rPr")
                .Where(properties => properties.Ancestor(C, "title").IsNone)
                .Concat(owner.Descendants(A, "defRPr")
                    .Where(properties => properties.Ancestor(C, "title").IsNone));
    }
}
