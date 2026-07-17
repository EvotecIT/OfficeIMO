using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTable {

        /// <summary>
        ///     Retrieves a cell at the specified row and column index.
        /// </summary>
        /// <param name="row">Zero-based row index.</param>
        /// <param name="column">Zero-based column index.</param>
        public PowerPointTableCell GetCell(int row, int column) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            A.TableRow tableRow = table.Elements<A.TableRow>().ElementAt(row);
            A.TableCell cell = tableRow.Elements<A.TableCell>().ElementAt(column);
            return new PowerPointTableCell(cell, _slidePart);
        }

        /// <summary>
        ///     Applies an action to all cells in the table.
        /// </summary>
        public void ApplyToCells(Action<PowerPointTableCell> action) {
            if (action == null) {
                throw new ArgumentNullException(nameof(action));
            }
            if (Rows == 0 || Columns == 0) {
                return;
            }

            ApplyToCells(0, Rows - 1, 0, Columns - 1, action);
        }

        /// <summary>
        ///     Applies an action to a rectangular range of cells.
        /// </summary>
        public void ApplyToCells(int startRow, int endRow, int startColumn, int endColumn, Action<PowerPointTableCell> action) {
            if (action == null) {
                throw new ArgumentNullException(nameof(action));
            }
            if (Rows == 0 || Columns == 0) {
                return;
            }

            int topRow = Math.Min(startRow, endRow);
            int bottomRow = Math.Max(startRow, endRow);
            int leftColumn = Math.Min(startColumn, endColumn);
            int rightColumn = Math.Max(startColumn, endColumn);

            if (topRow < 0 || leftColumn < 0) {
                throw new ArgumentOutOfRangeException("Row and column indices must be non-negative.");
            }
            if (bottomRow >= Rows || rightColumn >= Columns) {
                throw new ArgumentOutOfRangeException("Range exceeds table bounds.");
            }

            for (int r = topRow; r <= bottomRow; r++) {
                for (int c = leftColumn; c <= rightColumn; c++) {
                    action(GetCell(r, c));
                }
            }
        }

        /// <summary>
        ///     Applies an action to a specific row.
        /// </summary>
        public void ApplyToRow(int rowIndex, Action<PowerPointTableCell> action) {
            if (rowIndex < 0 || rowIndex >= Rows) {
                throw new ArgumentOutOfRangeException(nameof(rowIndex));
            }
            ApplyToCells(rowIndex, rowIndex, 0, Columns - 1, action);
        }

        /// <summary>
        ///     Applies an action to a specific column.
        /// </summary>
        public void ApplyToColumn(int columnIndex, Action<PowerPointTableCell> action) {
            if (columnIndex < 0 || columnIndex >= Columns) {
                throw new ArgumentOutOfRangeException(nameof(columnIndex));
            }
            ApplyToCells(0, Rows - 1, columnIndex, columnIndex, action);
        }

        /// <summary>
        ///     Sets cell padding in points for all cells.
        /// </summary>
        public void SetCellPaddingPoints(double? left, double? top, double? right, double? bottom) {
            ApplyToCells(cell => {
                cell.PaddingLeftPoints = left;
                cell.PaddingTopPoints = top;
                cell.PaddingRightPoints = right;
                cell.PaddingBottomPoints = bottom;
            });
        }

        /// <summary>
        ///     Sets cell padding in points for a range of cells.
        /// </summary>
        public void SetCellPaddingPoints(double? left, double? top, double? right, double? bottom,
            int startRow, int endRow, int startColumn, int endColumn) {
            ApplyToCells(startRow, endRow, startColumn, endColumn, cell => {
                cell.PaddingLeftPoints = left;
                cell.PaddingTopPoints = top;
                cell.PaddingRightPoints = right;
                cell.PaddingBottomPoints = bottom;
            });
        }

        /// <summary>
        ///     Sets cell padding in centimeters for all cells.
        /// </summary>
        public void SetCellPaddingCm(double? leftCm, double? topCm, double? rightCm, double? bottomCm) {
            ApplyToCells(cell => {
                cell.PaddingLeftCm = leftCm;
                cell.PaddingTopCm = topCm;
                cell.PaddingRightCm = rightCm;
                cell.PaddingBottomCm = bottomCm;
            });
        }

        /// <summary>
        ///     Sets cell padding in inches for all cells.
        /// </summary>
        public void SetCellPaddingInches(double? leftInches, double? topInches, double? rightInches, double? bottomInches) {
            ApplyToCells(cell => {
                cell.PaddingLeftInches = leftInches;
                cell.PaddingTopInches = topInches;
                cell.PaddingRightInches = rightInches;
                cell.PaddingBottomInches = bottomInches;
            });
        }

        /// <summary>
        ///     Sets cell alignment for all cells.
        /// </summary>
        public void SetCellAlignment(A.TextAlignmentTypeValues? horizontal, A.TextAnchoringTypeValues? vertical) {
            ApplyToCells(cell => {
                cell.HorizontalAlignment = horizontal;
                cell.VerticalAlignment = vertical;
            });
        }

        /// <summary>
        ///     Sets cell alignment for a range of cells.
        /// </summary>
        public void SetCellAlignment(A.TextAlignmentTypeValues? horizontal, A.TextAnchoringTypeValues? vertical,
            int startRow, int endRow, int startColumn, int endColumn) {
            ApplyToCells(startRow, endRow, startColumn, endColumn, cell => {
                cell.HorizontalAlignment = horizontal;
                cell.VerticalAlignment = vertical;
            });
        }

        /// <summary>
        ///     Applies borders to all cells.
        /// </summary>
        public void SetCellBorders(TableCellBorders borders, string color, double? widthPoints = null) {
            ApplyToCells(cell => cell.SetBorders(borders, color, widthPoints));
        }

        /// <summary>
        ///     Applies dashed borders to all cells.
        /// </summary>
        public void SetCellBorders(TableCellBorders borders, string color, double? widthPoints, A.PresetLineDashValues dash) {
            ApplyToCells(cell => cell.SetBorders(borders, color, widthPoints, dash));
        }

        /// <summary>
        ///     Clears borders for all cells.
        /// </summary>
        public void ClearCellBorders(TableCellBorders borders) {
            ApplyToCells(cell => cell.ClearBorders(borders));
        }

        /// <summary>
        ///     Merges a rectangular range of cells into the top-left cell.
        /// </summary>
        /// <param name="startRow">Zero-based start row.</param>
        /// <param name="startColumn">Zero-based start column.</param>
        /// <param name="endRow">Zero-based end row.</param>
        /// <param name="endColumn">Zero-based end column.</param>
        /// <param name="clearMergedContent">Whether to clear text from merged cells.</param>
        public void MergeCells(int startRow, int startColumn, int endRow, int endColumn, bool clearMergedContent = true) {
            int topRow = Math.Min(startRow, endRow);
            int bottomRow = Math.Max(startRow, endRow);
            int leftColumn = Math.Min(startColumn, endColumn);
            int rightColumn = Math.Max(startColumn, endColumn);

            if (topRow < 0 || leftColumn < 0) {
                throw new ArgumentOutOfRangeException("Row and column indices must be non-negative.");
            }
            if (bottomRow >= Rows || rightColumn >= Columns) {
                throw new ArgumentOutOfRangeException("Merge range exceeds table bounds.");
            }

            int rowSpan = bottomRow - topRow + 1;
            int colSpan = rightColumn - leftColumn + 1;
            if (rowSpan == 1 && colSpan == 1) {
                return;
            }

            A.Table table = TableElement;
            for (int r = topRow; r <= bottomRow; r++) {
                A.TableRow row = table.Elements<A.TableRow>().ElementAt(r);
                for (int c = leftColumn; c <= rightColumn; c++) {
                    A.TableCell cell = row.Elements<A.TableCell>().ElementAt(c);
                    bool isAnchor = r == topRow && c == leftColumn;

                    if (isAnchor) {
                        cell.RowSpan = rowSpan > 1 ? rowSpan : null;
                        cell.GridSpan = colSpan > 1 ? colSpan : null;
                        cell.HorizontalMerge = null;
                        cell.VerticalMerge = null;
                        continue;
                    }

                    cell.RowSpan = null;
                    cell.GridSpan = null;
                    cell.HorizontalMerge = c > leftColumn ? true : (bool?)null;
                    cell.VerticalMerge = r > topRow ? true : (bool?)null;

                    if (clearMergedContent) {
                        ClearMergedCellText(cell);
                    }
                }
            }
        }

        private void ClearCellText(A.TableCell cell) {
            if (cell.TextBody == null) {
                return;
            }

            string[] discardedSoundIds = PowerPointEmbeddedSound
                .GetRelationshipIds(cell.TextBody);
            cell.TextBody.RemoveAllChildren<A.Paragraph>();
            cell.TextBody.Append(new A.Paragraph(new A.Run(new A.Text(string.Empty))));
            PowerPointEmbeddedSound.RemoveIfUnused(_slidePart,
                discardedSoundIds);
        }

        private void ClearMergedCellText(A.TableCell cell) {
            ClearCellText(cell);
        }
    }
}
