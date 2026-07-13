using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private const string TableStylesResourceName = "OfficeIMO.PowerPoint.Resources.tableStyles.xml";

        private enum TableCellBorderEdge {
            Left,
            Top,
            Right,
            Bottom
        }

        private static A.TableStyleEntry? ResolveTableStyle(PowerPointTable table) {
            string? styleId = table.StyleId;
            if (string.IsNullOrWhiteSpace(styleId)) {
                return null;
            }

            string resolvedStyleId = styleId!;
            PresentationPart? presentationPart = table.OwnerSlide?.SlidePart
                .GetParentParts()
                .OfType<PresentationPart>()
                .FirstOrDefault();
            if (presentationPart != null) {
                if (presentationPart.TableStylesPart?.TableStyleList == null) {
                    PowerPointUtils.CreateTableStylesPart(presentationPart);
                }

                A.TableStyleEntry? packageStyle = FindTableStyle(presentationPart.TableStylesPart?.TableStyleList, resolvedStyleId);
                if (packageStyle != null) {
                    return packageStyle;
                }
            }

            using Stream? resource = typeof(PowerPointSlideImageRenderer).Assembly.GetManifestResourceStream(TableStylesResourceName);
            if (resource == null) {
                return null;
            }

            XDocument document = PowerPointXmlReader.LoadPackagePartXml(resource);
            string? xml = document.Root?.ToString(SaveOptions.DisableFormatting);
            if (string.IsNullOrWhiteSpace(xml)) {
                return null;
            }

            string resolvedXml = xml!;
            return FindTableStyle(new A.TableStyleList(resolvedXml), resolvedStyleId);
        }

        private static A.TableStyleEntry? FindTableStyle(A.TableStyleList? styleList, string styleId) {
            if (styleList == null) {
                return null;
            }

            return styleList.Elements<A.TableStyleEntry>()
                .FirstOrDefault(style => string.Equals(style.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase));
        }

        private static OfficeColor? ResolveTableStyleFillColor(PowerPointTable table, int row, int column, A.TableStyleEntry? tableStyle, A.ColorScheme? colorScheme) {
            foreach (A.TablePartStyleType region in EnumerateApplicableTableStyleRegions(table, row, column, tableStyle)) {
                A.SolidFill? solidFill = region.TableCellStyle?.GetFirstChild<A.FillProperties>()?.SolidFill;
                OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(solidFill, colorScheme);
                if (color.HasValue) {
                    return color;
                }
            }

            return null;
        }

        private static OfficeBorderSide? ResolveTableStyleBorderSide(PowerPointTable table, int row, int column, int rowSpan, int columnSpan, A.TableStyleEntry? tableStyle, TableCellBorderEdge edge, A.ColorScheme? colorScheme) {
            foreach (A.TablePartStyleType region in EnumerateApplicableTableStyleRegions(table, row, column, tableStyle)) {
                A.TableCellBorders? borders = region.TableCellStyle?.TableCellBorders;
                A.Outline? outline = ResolveTableStyleBorderOutline(table, row, column, rowSpan, columnSpan, borders, edge);
                if (outline != null) {
                    return ResolveTableCellBorderSide(outline, colorScheme);
                }
            }

            return null;
        }

        private static A.Outline? ResolveTableStyleBorderOutline(PowerPointTable table, int row, int column, int rowSpan, int columnSpan, A.TableCellBorders? borders, TableCellBorderEdge edge) {
            if (borders == null) {
                return null;
            }

            bool firstRow = row == 0;
            bool lastRow = row + Math.Max(1, rowSpan) >= table.Rows;
            bool firstColumn = column == 0;
            bool lastColumn = column + Math.Max(1, columnSpan) >= table.Columns;
            return edge switch {
                TableCellBorderEdge.Left => firstColumn
                    ? borders.LeftBorder?.Outline
                    : borders.InsideVerticalBorder?.Outline ?? borders.LeftBorder?.Outline,
                TableCellBorderEdge.Top => firstRow
                    ? borders.TopBorder?.Outline
                    : borders.InsideHorizontalBorder?.Outline ?? borders.TopBorder?.Outline,
                TableCellBorderEdge.Right => lastColumn
                    ? borders.RightBorder?.Outline
                    : borders.InsideVerticalBorder?.Outline ?? borders.RightBorder?.Outline,
                TableCellBorderEdge.Bottom => lastRow
                    ? borders.BottomBorder?.Outline
                    : borders.InsideHorizontalBorder?.Outline ?? borders.BottomBorder?.Outline,
                _ => null
            };
        }

        private static IEnumerable<A.TablePartStyleType> EnumerateApplicableTableStyleRegions(PowerPointTable table, int row, int column, A.TableStyleEntry? tableStyle) {
            if (tableStyle == null) {
                yield break;
            }

            bool firstRow = row == 0;
            bool lastRow = row == table.Rows - 1;
            bool firstColumn = column == 0;
            bool lastColumn = column == table.Columns - 1;

            if (table.FirstRow && table.FirstColumn && firstRow && firstColumn && tableStyle.NorthwestCell != null) {
                yield return tableStyle.NorthwestCell;
            }

            if (table.FirstRow && table.LastColumn && firstRow && lastColumn && tableStyle.NortheastCell != null) {
                yield return tableStyle.NortheastCell;
            }

            if (table.LastRow && table.FirstColumn && lastRow && firstColumn && tableStyle.SouthwestCell != null) {
                yield return tableStyle.SouthwestCell;
            }

            if (table.LastRow && table.LastColumn && lastRow && lastColumn && tableStyle.SoutheastCell != null) {
                yield return tableStyle.SoutheastCell;
            }

            if (table.FirstRow && firstRow && tableStyle.FirstRow != null) {
                yield return tableStyle.FirstRow;
            }

            if (table.LastRow && lastRow && tableStyle.LastRow != null) {
                yield return tableStyle.LastRow;
            }

            if (table.FirstColumn && firstColumn && tableStyle.FirstColumn != null) {
                yield return tableStyle.FirstColumn;
            }

            if (table.LastColumn && lastColumn && tableStyle.LastColumn != null) {
                yield return tableStyle.LastColumn;
            }

            if (table.BandedRows && !IsExcludedTableStyleBandRow(table, row)) {
                A.TablePartStyleType? band = ResolveHorizontalBand(table, row, tableStyle);
                if (band != null) {
                    yield return band;
                }
            }

            if (table.BandedColumns && !IsExcludedTableStyleBandColumn(table, column)) {
                A.TablePartStyleType? band = ResolveVerticalBand(table, column, tableStyle);
                if (band != null) {
                    yield return band;
                }
            }

            if (tableStyle.WholeTable != null) {
                yield return tableStyle.WholeTable;
            }
        }

        private static bool IsExcludedTableStyleBandRow(PowerPointTable table, int row) =>
            (table.FirstRow && row == 0) || (table.LastRow && row == table.Rows - 1);

        private static bool IsExcludedTableStyleBandColumn(PowerPointTable table, int column) =>
            (table.FirstColumn && column == 0) || (table.LastColumn && column == table.Columns - 1);

        private static A.TablePartStyleType? ResolveHorizontalBand(PowerPointTable table, int row, A.TableStyleEntry tableStyle) {
            int visualRow = row - (table.FirstRow ? 1 : 0);
            return visualRow % 2 == 0
                ? tableStyle.Band1Horizontal
                : tableStyle.Band2Horizontal;
        }

        private static A.TablePartStyleType? ResolveVerticalBand(PowerPointTable table, int column, A.TableStyleEntry tableStyle) {
            int visualColumn = column - (table.FirstColumn ? 1 : 0);
            return visualColumn % 2 == 0
                ? tableStyle.Band1Vertical
                : tableStyle.Band2Vertical;
        }
    }
}
