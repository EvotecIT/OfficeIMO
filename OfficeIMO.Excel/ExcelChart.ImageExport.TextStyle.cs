using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        private static bool TryGetImageExportTitleColor(C.Chart? chart, WorkbookPart workbookPart, out OfficeColor color) {
            color = default;
            C.Title? title = chart?.GetFirstChild<C.Title>();
            return title != null && TryGetFirstImageExportTextColor(new[] { title }, workbookPart, out color);
        }

        private static bool TryGetImageExportTitleFontFamily(C.Chart? chart, out string? fontFamily) {
            C.Title? title = chart?.GetFirstChild<C.Title>();
            if (title != null) {
                return TryGetFirstImageExportTextFontFamily(new[] { title }, out fontFamily);
            }

            fontFamily = default;
            return false;
        }

        private static bool TryGetImageExportTitleFontSize(C.Chart? chart, out double fontSize) {
            C.Title? title = chart?.GetFirstChild<C.Title>();
            if (title != null) {
                return TryGetFirstImageExportTextFontSize(new[] { title }, out fontSize);
            }

            fontSize = default;
            return false;
        }

        private static bool TryGetImageExportTitleFontStyle(C.Chart? chart, out OfficeFontStyle fontStyle) {
            C.Title? title = chart?.GetFirstChild<C.Title>();
            if (title != null) {
                return TryGetFirstImageExportTextFontStyle(new[] { title }, out fontStyle);
            }

            fontStyle = default;
            return false;
        }

        private static bool TryGetImageExportLegendTextColor(C.Chart? chart, WorkbookPart workbookPart, out OfficeColor color) =>
            TryGetFirstImageExportTextColor(GetImageExportLegendTextOwners(chart), workbookPart, out color);

        private static bool TryGetImageExportDataLabelTextColor(C.PlotArea? plotArea, WorkbookPart workbookPart, out OfficeColor color) =>
            TryGetFirstImageExportTextColor(GetImageExportDataLabelTextOwners(plotArea), workbookPart, out color);

        private static bool TryGetImageExportDataLabelFillColor(C.PlotArea? plotArea, WorkbookPart workbookPart, out OfficeColor color) =>
            TryGetFirstImageExportShapeFill(GetImageExportDataLabelTextOwners(plotArea), workbookPart, out color);

        private static bool TryGetImageExportDataLabelLineColor(C.PlotArea? plotArea, WorkbookPart workbookPart, out OfficeColor color) =>
            TryGetFirstImageExportShapeLine(GetImageExportDataLabelTextOwners(plotArea), workbookPart, out color);

        private static bool TryGetImageExportDataLabelLineWidth(C.PlotArea? plotArea, out double width) =>
            TryGetFirstImageExportShapeLineWidth(GetImageExportDataLabelTextOwners(plotArea), out width);

        private static bool TryGetImageExportDataLabelLineDashStyle(C.PlotArea? plotArea, out OfficeStrokeDashStyle dashStyle) =>
            TryGetFirstImageExportShapeLineDashStyle(GetImageExportDataLabelTextOwners(plotArea), out dashStyle);

        private static bool TryGetImageExportAxisTextColor(C.PlotArea? plotArea, WorkbookPart workbookPart, out OfficeColor color) =>
            TryGetFirstImageExportAxisLabelTextColor(plotArea, workbookPart, out color);

        private static bool TryGetImageExportAxisTitleTextColor(C.PlotArea? plotArea, WorkbookPart workbookPart, out OfficeColor color) =>
            TryGetFirstImageExportTextColor(GetImageExportAxisTitleTextOwners(plotArea), workbookPart, out color);

        private static bool TryGetImageExportLegendFontSize(C.Chart? chart, out double fontSize) =>
            TryGetFirstImageExportTextFontSize(GetImageExportLegendTextOwners(chart), out fontSize);

        private static bool TryGetImageExportLegendFontFamily(C.Chart? chart, out string? fontFamily) =>
            TryGetFirstImageExportTextFontFamily(GetImageExportLegendTextOwners(chart), out fontFamily);

        private static bool TryGetImageExportLegendFontStyle(C.Chart? chart, out OfficeFontStyle fontStyle) =>
            TryGetFirstImageExportTextFontStyle(GetImageExportLegendTextOwners(chart), out fontStyle);

        private static bool TryGetImageExportDataLabelFontSize(C.PlotArea? plotArea, out double fontSize) =>
            TryGetFirstImageExportTextFontSize(GetImageExportDataLabelTextOwners(plotArea), out fontSize);

        private static bool TryGetImageExportDataLabelFontFamily(C.PlotArea? plotArea, out string? fontFamily) =>
            TryGetFirstImageExportTextFontFamily(GetImageExportDataLabelTextOwners(plotArea), out fontFamily);

        private static bool TryGetImageExportDataLabelFontStyle(C.PlotArea? plotArea, out OfficeFontStyle fontStyle) =>
            TryGetFirstImageExportTextFontStyle(GetImageExportDataLabelTextOwners(plotArea), out fontStyle);

        private static bool TryGetImageExportAxisLabelFontSize(C.PlotArea? plotArea, out double fontSize) =>
            TryGetFirstImageExportAxisLabelTextFontSize(plotArea, out fontSize);

        private static bool TryGetImageExportAxisTitleFontSize(C.PlotArea? plotArea, out double fontSize) =>
            TryGetFirstImageExportTextFontSize(GetImageExportAxisTitleTextOwners(plotArea), out fontSize);

        private static bool TryGetImageExportAxisTextFontFamily(C.PlotArea? plotArea, out string? fontFamily) =>
            TryGetFirstImageExportAxisLabelTextFontFamily(plotArea, out fontFamily);

        private static bool TryGetImageExportAxisTitleFontFamily(C.PlotArea? plotArea, out string? fontFamily) =>
            TryGetFirstImageExportTextFontFamily(GetImageExportAxisTitleTextOwners(plotArea), out fontFamily);

        private static bool TryGetImageExportAxisTextFontStyle(C.PlotArea? plotArea, out OfficeFontStyle fontStyle) =>
            TryGetFirstImageExportAxisLabelTextFontStyle(plotArea, out fontStyle);

        private static bool TryGetImageExportAxisTitleFontStyle(C.PlotArea? plotArea, out OfficeFontStyle fontStyle) =>
            TryGetFirstImageExportTextFontStyle(GetImageExportAxisTitleTextOwners(plotArea), out fontStyle);

        private static bool HasUnsupportedImageExportTextStyle(C.ChartSpace chartSpace, WorkbookPart workbookPart) {
            C.Chart? chart = chartSpace.GetFirstChild<C.Chart>();
            C.PlotArea? plotArea = chart?.GetFirstChild<C.PlotArea>();
            C.Title? chartTitle = chart?.GetFirstChild<C.Title>();
            if (chartTitle != null && HasUnsupportedImageExportTextStyle(chartTitle, allowSolidFill: true, allowFontFamily: true, allowFontSize: true, allowFontStyle: true, workbookPart)) {
                return true;
            }

            IReadOnlyList<OpenXmlCompositeElement> bodyOwners = GetImageExportBodyTextOwners(chart, plotArea).ToArray();
            foreach (OpenXmlCompositeElement owner in bodyOwners) {
                if (HasUnsupportedImageExportTextStyle(owner, allowSolidFill: true, allowFontFamily: true, allowFontSize: true, allowFontStyle: true, workbookPart)) {
                    return true;
                }
            }

            IReadOnlyList<OpenXmlCompositeElement> axisLabelOwners = GetImageExportAxisLabelTextOwners(plotArea).ToArray();
            foreach (OpenXmlCompositeElement owner in axisLabelOwners) {
                if (HasUnsupportedImageExportAxisLabelTextStyle(owner, allowSolidFill: true, allowFontFamily: true, allowFontSize: true, allowFontStyle: true, workbookPart)) {
                    return true;
                }
            }

            IReadOnlyList<OpenXmlCompositeElement> axisTitleOwners = GetImageExportAxisTitleTextOwners(plotArea).ToArray();
            foreach (OpenXmlCompositeElement owner in axisTitleOwners) {
                if (HasUnsupportedImageExportTextStyle(owner, allowSolidFill: true, allowFontFamily: true, allowFontSize: true, allowFontStyle: true, workbookPart)) {
                    return true;
                }
            }

            IReadOnlyList<OpenXmlCompositeElement> legendOwners = GetImageExportLegendTextOwners(chart).ToArray();
            IReadOnlyList<OpenXmlCompositeElement> dataLabelOwners = GetImageExportDataLabelTextOwners(plotArea).ToArray();
            return HasConflictingImageExportTextColors(legendOwners, workbookPart) ||
                HasConflictingImageExportTextColors(dataLabelOwners, workbookPart) ||
                HasConflictingImageExportAxisLabelTextColors(axisLabelOwners, workbookPart) ||
                HasConflictingImageExportTextColors(axisTitleOwners, workbookPart) ||
                HasConflictingImageExportTextFontFamilies(chartTitle == null ? Array.Empty<OpenXmlCompositeElement>() : new[] { chartTitle }) ||
                HasConflictingImageExportTextFontFamilies(legendOwners) ||
                HasConflictingImageExportTextFontFamilies(dataLabelOwners) ||
                HasConflictingImageExportAxisLabelTextFontFamilies(axisLabelOwners) ||
                HasConflictingImageExportTextFontFamilies(axisTitleOwners) ||
                HasConflictingImageExportTextFontSizes(chartTitle == null ? Array.Empty<OpenXmlCompositeElement>() : new[] { chartTitle }) ||
                HasConflictingImageExportTextFontStyles(chartTitle == null ? Array.Empty<OpenXmlCompositeElement>() : new[] { chartTitle }) ||
                HasConflictingImageExportTextFontSizes(dataLabelOwners) ||
                HasConflictingImageExportAxisLabelTextFontSizes(axisLabelOwners) ||
                HasConflictingImageExportTextFontSizes(axisTitleOwners) ||
                HasConflictingImageExportTextFontStyles(legendOwners) ||
                HasConflictingImageExportTextFontStyles(dataLabelOwners) ||
                HasConflictingImageExportAxisLabelTextFontStyles(axisLabelOwners) ||
                HasConflictingImageExportTextFontStyles(axisTitleOwners);
        }

        private static IEnumerable<OpenXmlCompositeElement> GetImageExportLegendTextOwners(C.Chart? chart) {
            C.Legend? legend = chart?.GetFirstChild<C.Legend>();
            if (legend != null) {
                yield return legend;
            }
        }

        private static IEnumerable<OpenXmlCompositeElement> GetImageExportDataLabelTextOwners(C.PlotArea? plotArea) {
            if (plotArea == null) {
                yield break;
            }

            foreach (C.DataLabels labels in plotArea.Descendants<C.DataLabels>()) {
                yield return labels;
            }
        }

        private static IEnumerable<OpenXmlCompositeElement> GetImageExportBodyTextOwners(C.Chart? chart, C.PlotArea? plotArea) {
            foreach (OpenXmlCompositeElement legend in GetImageExportLegendTextOwners(chart)) {
                yield return legend;
            }

            foreach (OpenXmlCompositeElement labels in GetImageExportDataLabelTextOwners(plotArea)) {
                yield return labels;
            }
        }

        private static bool TryGetFirstImageExportShapeFill(IEnumerable<OpenXmlCompositeElement> owners, WorkbookPart workbookPart, out OfficeColor color) {
            foreach (OpenXmlCompositeElement owner in owners) {
                if (TryGetSolidFill(owner.GetFirstChild<C.ChartShapeProperties>(), workbookPart, out color)) {
                    return true;
                }
            }

            color = default;
            return false;
        }

        private static bool TryGetFirstImageExportShapeLine(IEnumerable<OpenXmlCompositeElement> owners, WorkbookPart workbookPart, out OfficeColor color) {
            foreach (OpenXmlCompositeElement owner in owners) {
                if (TryGetSolidLine(owner.GetFirstChild<C.ChartShapeProperties>(), workbookPart, out color)) {
                    return true;
                }
            }

            color = default;
            return false;
        }

        private static bool TryGetFirstImageExportShapeLineWidth(IEnumerable<OpenXmlCompositeElement> owners, out double width) {
            foreach (OpenXmlCompositeElement owner in owners) {
                if (TryGetLineWidth(owner.GetFirstChild<C.ChartShapeProperties>(), out width)) {
                    return true;
                }
            }

            width = default;
            return false;
        }

        private static bool TryGetFirstImageExportShapeLineDashStyle(IEnumerable<OpenXmlCompositeElement> owners, out OfficeStrokeDashStyle dashStyle) {
            foreach (OpenXmlCompositeElement owner in owners) {
                if (TryGetLineDashStyle(owner.GetFirstChild<C.ChartShapeProperties>(), out dashStyle)) {
                    return true;
                }
            }

            dashStyle = default;
            return false;
        }

        private static IEnumerable<OpenXmlCompositeElement> GetImageExportAxisTextOwners(C.PlotArea? plotArea) {
            foreach (OpenXmlCompositeElement axis in GetImageExportAxisLabelTextOwners(plotArea)) {
                yield return axis;
            }

            foreach (OpenXmlCompositeElement title in GetImageExportAxisTitleTextOwners(plotArea)) {
                yield return title;
            }
        }

        private static IEnumerable<OpenXmlCompositeElement> GetImageExportAxisLabelTextOwners(C.PlotArea? plotArea) {
            if (plotArea == null) {
                yield break;
            }

            foreach (OpenXmlCompositeElement axis in GetImageExportAxes(plotArea)) {
                yield return axis;
            }
        }

        private static IEnumerable<OpenXmlCompositeElement> GetImageExportAxisTitleTextOwners(C.PlotArea? plotArea) {
            if (plotArea == null) {
                yield break;
            }

            foreach (OpenXmlCompositeElement axis in GetImageExportAxes(plotArea)) {
                C.Title? title = axis.GetFirstChild<C.Title>();
                if (title != null) {
                    yield return title;
                }
            }
        }

        private static bool TryGetFirstImageExportTextColor(IEnumerable<OpenXmlCompositeElement> owners, WorkbookPart workbookPart, out OfficeColor color) {
            foreach (OpenXmlCompositeElement owner in owners) {
                foreach (OpenXmlCompositeElement properties in GetImageExportTextRunProperties(owner)) {
                    if (TryGetSolidFill(properties, workbookPart, out color)) {
                        return true;
                    }
                }
            }

            color = default;
            return false;
        }

        private static bool TryGetFirstImageExportAxisLabelTextColor(C.PlotArea? plotArea, WorkbookPart workbookPart, out OfficeColor color) {
            if (plotArea != null) {
                foreach (OpenXmlCompositeElement axis in GetImageExportAxes(plotArea)) {
                    foreach (OpenXmlCompositeElement properties in GetImageExportAxisLabelTextRunProperties(axis)) {
                        if (TryGetSolidFill(properties, workbookPart, out color)) {
                            return true;
                        }
                    }
                }
            }

            color = default;
            return false;
        }

        private static bool TryGetFirstImageExportTextFontFamily(IEnumerable<OpenXmlCompositeElement> owners, out string? fontFamily) {
            foreach (OpenXmlCompositeElement owner in owners) {
                foreach (OpenXmlCompositeElement properties in GetImageExportTextRunProperties(owner)) {
                    if (TryGetImageExportTextFontFamily(properties, out fontFamily)) {
                        return true;
                    }
                }
            }

            fontFamily = default;
            return false;
        }

        private static bool TryGetFirstImageExportTextFontSize(IEnumerable<OpenXmlCompositeElement> owners, out double fontSize) {
            foreach (OpenXmlCompositeElement owner in owners) {
                foreach (OpenXmlCompositeElement properties in GetImageExportTextRunProperties(owner)) {
                    if (TryGetImageExportTextFontSize(properties, out fontSize)) {
                        return true;
                    }
                }
            }

            fontSize = default;
            return false;
        }

        private static bool TryGetFirstImageExportAxisLabelTextFontSize(C.PlotArea? plotArea, out double fontSize) {
            if (plotArea != null) {
                foreach (OpenXmlCompositeElement axis in GetImageExportAxes(plotArea)) {
                    foreach (OpenXmlCompositeElement properties in GetImageExportAxisLabelTextRunProperties(axis)) {
                        if (TryGetImageExportTextFontSize(properties, out fontSize)) {
                            return true;
                        }
                    }
                }
            }

            fontSize = default;
            return false;
        }

        private static bool TryGetFirstImageExportAxisLabelTextFontFamily(C.PlotArea? plotArea, out string? fontFamily) {
            if (plotArea != null) {
                foreach (OpenXmlCompositeElement axis in GetImageExportAxes(plotArea)) {
                    foreach (OpenXmlCompositeElement properties in GetImageExportAxisLabelTextRunProperties(axis)) {
                        if (TryGetImageExportTextFontFamily(properties, out fontFamily)) {
                            return true;
                        }
                    }
                }
            }

            fontFamily = default;
            return false;
        }

        private static bool TryGetFirstImageExportTextFontStyle(IEnumerable<OpenXmlCompositeElement> owners, out OfficeFontStyle fontStyle) {
            foreach (OpenXmlCompositeElement owner in owners) {
                foreach (OpenXmlCompositeElement properties in GetImageExportTextRunProperties(owner)) {
                    if (TryGetImageExportTextFontStyle(properties, out fontStyle)) {
                        return true;
                    }
                }
            }

            fontStyle = default;
            return false;
        }

        private static bool TryGetFirstImageExportAxisLabelTextFontStyle(C.PlotArea? plotArea, out OfficeFontStyle fontStyle) {
            if (plotArea != null) {
                foreach (OpenXmlCompositeElement axis in GetImageExportAxes(plotArea)) {
                    foreach (OpenXmlCompositeElement properties in GetImageExportAxisLabelTextRunProperties(axis)) {
                        if (TryGetImageExportTextFontStyle(properties, out fontStyle)) {
                            return true;
                        }
                    }
                }
            }

            fontStyle = default;
            return false;
        }

        private static bool HasConflictingImageExportTextColors(IEnumerable<OpenXmlCompositeElement> owners, WorkbookPart workbookPart) {
            string? first = null;
            foreach (OpenXmlCompositeElement owner in owners) {
                foreach (OpenXmlCompositeElement properties in GetImageExportTextRunProperties(owner)) {
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

        private static bool HasConflictingImageExportAxisLabelTextColors(IEnumerable<OpenXmlCompositeElement> owners, WorkbookPart workbookPart) {
            string? first = null;
            foreach (OpenXmlCompositeElement owner in owners) {
                foreach (OpenXmlCompositeElement properties in GetImageExportAxisLabelTextRunProperties(owner)) {
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

        private static bool HasConflictingImageExportTextFontFamilies(IEnumerable<OpenXmlCompositeElement> owners) {
            string? first = null;
            foreach (OpenXmlCompositeElement owner in owners) {
                foreach (OpenXmlCompositeElement properties in GetImageExportTextRunProperties(owner)) {
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

        private static bool HasConflictingImageExportTextFontSizes(IEnumerable<OpenXmlCompositeElement> owners) {
            double? first = null;
            foreach (OpenXmlCompositeElement owner in owners) {
                foreach (OpenXmlCompositeElement properties in GetImageExportTextRunProperties(owner)) {
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

        private static bool HasConflictingImageExportAxisLabelTextFontFamilies(IEnumerable<OpenXmlCompositeElement> owners) {
            string? first = null;
            foreach (OpenXmlCompositeElement owner in owners) {
                foreach (OpenXmlCompositeElement properties in GetImageExportAxisLabelTextRunProperties(owner)) {
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

        private static bool HasConflictingImageExportAxisLabelTextFontSizes(IEnumerable<OpenXmlCompositeElement> owners) {
            double? first = null;
            foreach (OpenXmlCompositeElement owner in owners) {
                foreach (OpenXmlCompositeElement properties in GetImageExportAxisLabelTextRunProperties(owner)) {
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

        private static bool HasConflictingImageExportAxisLabelTextFontStyles(IEnumerable<OpenXmlCompositeElement> owners) {
            OfficeFontStyle? first = null;
            foreach (OpenXmlCompositeElement owner in owners) {
                foreach (OpenXmlCompositeElement properties in GetImageExportAxisLabelTextRunProperties(owner)) {
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

        private static bool HasConflictingImageExportTextFontStyles(IEnumerable<OpenXmlCompositeElement> owners) {
            OfficeFontStyle? first = null;
            foreach (OpenXmlCompositeElement owner in owners) {
                foreach (OpenXmlCompositeElement properties in GetImageExportTextRunProperties(owner)) {
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

        private static bool HasUnsupportedImageExportTextStyle(OpenXmlCompositeElement owner, bool allowSolidFill, bool allowFontFamily, bool allowFontSize, bool allowFontStyle, WorkbookPart workbookPart) {
            foreach (OpenXmlCompositeElement properties in GetImageExportTextRunProperties(owner)) {
                if (!IsSimpleSupportedImageExportTextProperties(properties, allowSolidFill, allowFontFamily, allowFontSize, allowFontStyle, workbookPart)) {
                    return true;
                }
            }

            return false;
        }

        private static bool HasUnsupportedImageExportAxisLabelTextStyle(OpenXmlCompositeElement owner, bool allowSolidFill, bool allowFontFamily, bool allowFontSize, bool allowFontStyle, WorkbookPart workbookPart) {
            foreach (OpenXmlCompositeElement properties in GetImageExportAxisLabelTextRunProperties(owner)) {
                if (!IsSimpleSupportedImageExportTextProperties(properties, allowSolidFill, allowFontFamily, allowFontSize, allowFontStyle, workbookPart)) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsSimpleSupportedImageExportTextProperties(OpenXmlCompositeElement properties, bool allowSolidFill, bool allowFontFamily, bool allowFontSize, bool allowFontStyle, WorkbookPart workbookPart) {
            if (properties is A.TextCharacterPropertiesType textProperties &&
                (textProperties.Bold != null || textProperties.Italic != null) &&
                !allowFontStyle) {
                return false;
            }
            if (properties is A.TextCharacterPropertiesType fontProperties && fontProperties.FontSize != null) {
                if (!allowFontSize || !TryGetImageExportTextFontSize(properties, out _)) {
                    return false;
                }
            }

            foreach (OpenXmlElement child in properties.ChildElements) {
                if (allowSolidFill && child is A.SolidFill) {
                    if (!TryGetSolidFill(properties, workbookPart, out _)) {
                        return false;
                    }

                    continue;
                }

                if (allowFontFamily && child is A.LatinFont) {
                    if (!TryGetImageExportTextFontFamily(properties, out _)) {
                        return false;
                    }

                    continue;
                }

                return false;
            }

            return true;
        }

        private static bool TryGetImageExportTextFontFamily(OpenXmlCompositeElement properties, out string? fontFamily) {
            string? typeface = properties.GetFirstChild<A.LatinFont>()?.Typeface?.Value;
            if (!string.IsNullOrWhiteSpace(typeface)) {
                fontFamily = typeface;
                return true;
            }

            fontFamily = default;
            return false;
        }

        private static bool IsThemeFontFamilyPlaceholder(string fontFamily) =>
            fontFamily.Length > 0 && fontFamily[0] == '+';

        private static bool TryGetImageExportTextFontSize(OpenXmlCompositeElement properties, out double fontSize) {
            if (properties is A.TextCharacterPropertiesType textProperties &&
                textProperties.FontSize != null &&
                textProperties.FontSize.Value > 0) {
                fontSize = textProperties.FontSize.Value / 100D;
                return true;
            }

            fontSize = default;
            return false;
        }

        private static bool TryGetImageExportTextFontStyle(OpenXmlCompositeElement properties, out OfficeFontStyle fontStyle) {
            fontStyle = OfficeFontStyle.Regular;
            if (properties is not A.TextCharacterPropertiesType textProperties) {
                return false;
            }

            bool hasStyle = false;
            if (textProperties.Bold != null) {
                hasStyle = true;
                if (textProperties.Bold.Value) {
                    fontStyle |= OfficeFontStyle.Bold;
                }
            }
            if (textProperties.Italic != null) {
                hasStyle = true;
                if (textProperties.Italic.Value) {
                    fontStyle |= OfficeFontStyle.Italic;
                }
            }

            return hasStyle;
        }

        private static IEnumerable<OpenXmlCompositeElement> GetImageExportTextRunProperties(OpenXmlCompositeElement owner) =>
            owner.Descendants<A.RunProperties>().Cast<OpenXmlCompositeElement>()
                .Concat(owner.Descendants<A.DefaultRunProperties>());

        private static IEnumerable<OpenXmlCompositeElement> GetImageExportAxisLabelTextRunProperties(OpenXmlCompositeElement owner) =>
            owner.Descendants<A.RunProperties>()
                .Where(properties => !properties.Ancestors<C.Title>().Any())
                .Cast<OpenXmlCompositeElement>()
                .Concat(owner.Descendants<A.DefaultRunProperties>()
                    .Where(properties => !properties.Ancestors<C.Title>().Any()));
    }
}
