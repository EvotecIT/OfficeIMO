using System;
using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Arranges shapes into a grid within their selection bounds.
        /// </summary>
        public void ArrangeShapesInGrid(IEnumerable<PowerPointShape> shapes, int columns, int rows,
            long gutterX = 0L, long gutterY = 0L, bool resizeToCell = true,
            PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            ArrangeShapesInGrid(shapes, GetSelectionBounds(shapes), columns, rows, gutterX, gutterY, resizeToCell, flow);
        }

        /// <summary>
        ///     Arranges shapes into a grid within a custom bounding box.
        /// </summary>
        public void ArrangeShapesInGrid(IEnumerable<PowerPointShape> shapes, PowerPointLayoutBox bounds, int columns,
            int rows, long gutterX = 0L, long gutterY = 0L, bool resizeToCell = true,
            PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }
            if (columns <= 0) {
                throw new ArgumentOutOfRangeException(nameof(columns));
            }
            if (rows <= 0) {
                throw new ArgumentOutOfRangeException(nameof(rows));
            }
            if (gutterX < 0) {
                throw new ArgumentOutOfRangeException(nameof(gutterX));
            }
            if (gutterY < 0) {
                throw new ArgumentOutOfRangeException(nameof(gutterY));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count == 0) {
                return;
            }

            PowerPointLayoutBox[,] grid = BuildGrid(bounds, columns, rows, gutterX, gutterY);
            int max = Math.Min(list.Count, columns * rows);

            for (int i = 0; i < max; i++) {
                (int row, int col) = flow == PowerPointShapeGridFlow.RowMajor
                    ? (i / columns, i % columns)
                    : (i % rows, i / rows);

                PowerPointLayoutBox cell = grid[row, col];
                PowerPointShape shape = list[i];
                if (resizeToCell) {
                    cell.ApplyTo(shape);
                } else {
                    shape.Left = cell.Left;
                    shape.Top = cell.Top;
                }
            }
        }

        /// <summary>
        ///     Arranges shapes into a grid that spans the slide bounds.        
        /// </summary>
        public void ArrangeShapesInGridToSlide(IEnumerable<PowerPointShape> shapes, int columns, int rows,
            long gutterX = 0L, long gutterY = 0L, bool resizeToCell = true,
            PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {  
            ArrangeShapesInGrid(shapes, GetSlideBounds(), columns, rows, gutterX, gutterY, resizeToCell, flow);
        }

        /// <summary>
        ///     Arranges shapes into a grid that spans the slide content bounds using a margin (EMUs).
        /// </summary>
        public void ArrangeShapesInGridToSlideContent(IEnumerable<PowerPointShape> shapes, int columns, int rows,
            long marginEmus, long gutterX = 0L, long gutterY = 0L, bool resizeToCell = true,
            PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            ArrangeShapesInGrid(shapes, GetSlideContentBounds(marginEmus), columns, rows, gutterX, gutterY,
                resizeToCell, flow);
        }

        /// <summary>
        ///     Arranges shapes into a grid that spans the slide content bounds using a margin (centimeters).
        /// </summary>
        public void ArrangeShapesInGridToSlideContentCm(IEnumerable<PowerPointShape> shapes, int columns, int rows,
            double marginCm, double gutterXCm = 0d, double gutterYCm = 0d, bool resizeToCell = true,
            PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            ArrangeShapesInGridToSlideContent(shapes, columns, rows,
                PowerPointUnits.FromCentimeters(marginCm),
                PowerPointUnits.FromCentimeters(gutterXCm),
                PowerPointUnits.FromCentimeters(gutterYCm),
                resizeToCell, flow);
        }

        /// <summary>
        ///     Arranges shapes into a grid that spans the slide content bounds using a margin (inches).
        /// </summary>
        public void ArrangeShapesInGridToSlideContentInches(IEnumerable<PowerPointShape> shapes, int columns, int rows,
            double marginInches, double gutterXInches = 0d, double gutterYInches = 0d, bool resizeToCell = true,
            PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            ArrangeShapesInGridToSlideContent(shapes, columns, rows,
                PowerPointUnits.FromInches(marginInches),
                PowerPointUnits.FromInches(gutterXInches),
                PowerPointUnits.FromInches(gutterYInches),
                resizeToCell, flow);
        }

        /// <summary>
        ///     Arranges shapes into a grid that spans the slide content bounds using a margin (points).
        /// </summary>
        public void ArrangeShapesInGridToSlideContentPoints(IEnumerable<PowerPointShape> shapes, int columns, int rows,
            double marginPoints, double gutterXPoints = 0d, double gutterYPoints = 0d, bool resizeToCell = true,
            PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            ArrangeShapesInGridToSlideContent(shapes, columns, rows,
                PowerPointUnits.FromPoints(marginPoints),
                PowerPointUnits.FromPoints(gutterXPoints),
                PowerPointUnits.FromPoints(gutterYPoints),
                resizeToCell, flow);
        }

        /// <summary>
        ///     Arranges shapes into an auto-sized grid within their selection bounds.
        /// </summary>
        public void ArrangeShapesInGridAuto(IEnumerable<PowerPointShape> shapes, long gutterX = 0L, long gutterY = 0L,
            bool resizeToCell = true, PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            ArrangeShapesInGridAuto(shapes, GetSelectionBounds(shapes),
                new PowerPointShapeGridOptions {
                    GutterX = gutterX,
                    GutterY = gutterY,
                    ResizeToCell = resizeToCell,
                    Flow = flow
                });
        }

        /// <summary>
        ///     Arranges shapes into an auto-sized grid within a custom bounding box.
        /// </summary>
        public void ArrangeShapesInGridAuto(IEnumerable<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            long gutterX = 0L, long gutterY = 0L, bool resizeToCell = true,
            PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            ArrangeShapesInGridAuto(shapes, bounds,
                new PowerPointShapeGridOptions {
                    GutterX = gutterX,
                    GutterY = gutterY,
                    ResizeToCell = resizeToCell,
                    Flow = flow
                });
        }

        /// <summary>
        ///     Arranges shapes into an auto-sized grid within a custom bounding box using options.
        /// </summary>
        public void ArrangeShapesInGridAuto(IEnumerable<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            PowerPointShapeGridOptions options) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count == 0) {
                return;
            }

            PowerPointShapeGridOptions resolvedOptions = options ?? new PowerPointShapeGridOptions();
            int columns = GetAutoGridColumns(list.Count, bounds, resolvedOptions.GutterX, resolvedOptions.GutterY,
                resolvedOptions.MinColumns, resolvedOptions.MaxColumns, resolvedOptions.TargetCellAspect);
            int rows = (int)Math.Ceiling(list.Count / (double)columns);

            ArrangeShapesInGrid(list, bounds, columns, rows, resolvedOptions.GutterX, resolvedOptions.GutterY,
                resolvedOptions.ResizeToCell, resolvedOptions.Flow);
        }

        /// <summary>
        ///     Arranges shapes into an auto-sized grid that spans the slide bounds.
        /// </summary>
        public void ArrangeShapesInGridAutoToSlide(IEnumerable<PowerPointShape> shapes, long gutterX = 0L, long gutterY = 0L,
            bool resizeToCell = true, PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            ArrangeShapesInGridAuto(shapes, GetSlideBounds(),
                new PowerPointShapeGridOptions {
                    GutterX = gutterX,
                    GutterY = gutterY,
                    ResizeToCell = resizeToCell,
                    Flow = flow
            });
        }

        /// <summary>
        ///     Arranges shapes into an auto-sized grid that spans the slide bounds using options.
        /// </summary>
        public void ArrangeShapesInGridAutoToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeGridOptions options) {
            ArrangeShapesInGridAuto(shapes, GetSlideBounds(), options);
        }

        /// <summary>
        ///     Arranges shapes into an auto-sized grid that spans the slide content bounds using a margin (EMUs).
        /// </summary>
        public void ArrangeShapesInGridAutoToSlideContent(IEnumerable<PowerPointShape> shapes, long marginEmus,
            long gutterX = 0L, long gutterY = 0L, bool resizeToCell = true,
            PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            ArrangeShapesInGridAuto(shapes, GetSlideContentBounds(marginEmus),
                new PowerPointShapeGridOptions {
                    GutterX = gutterX,
                    GutterY = gutterY,
                    ResizeToCell = resizeToCell,
                    Flow = flow
                });
        }

        /// <summary>
        ///     Arranges shapes into an auto-sized grid that spans the slide content bounds using a margin (EMUs) and options.
        /// </summary>
        public void ArrangeShapesInGridAutoToSlideContent(IEnumerable<PowerPointShape> shapes, long marginEmus,
            PowerPointShapeGridOptions options) {
            ArrangeShapesInGridAuto(shapes, GetSlideContentBounds(marginEmus), options);
        }

        /// <summary>
        ///     Arranges shapes into an auto-sized grid that spans the slide content bounds using a margin (centimeters).
        /// </summary>
        public void ArrangeShapesInGridAutoToSlideContentCm(IEnumerable<PowerPointShape> shapes, double marginCm,
            double gutterXCm = 0d, double gutterYCm = 0d, bool resizeToCell = true,
            PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            ArrangeShapesInGridAutoToSlideContent(shapes,
                PowerPointUnits.FromCentimeters(marginCm),
                PowerPointUnits.FromCentimeters(gutterXCm),
                PowerPointUnits.FromCentimeters(gutterYCm),
                resizeToCell, flow);
        }

        /// <summary>
        ///     Arranges shapes into an auto-sized grid that spans the slide content bounds using a margin (inches).
        /// </summary>
        public void ArrangeShapesInGridAutoToSlideContentInches(IEnumerable<PowerPointShape> shapes, double marginInches,
            double gutterXInches = 0d, double gutterYInches = 0d, bool resizeToCell = true,
            PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            ArrangeShapesInGridAutoToSlideContent(shapes,
                PowerPointUnits.FromInches(marginInches),
                PowerPointUnits.FromInches(gutterXInches),
                PowerPointUnits.FromInches(gutterYInches),
                resizeToCell, flow);
        }

        /// <summary>
        ///     Arranges shapes into an auto-sized grid that spans the slide content bounds using a margin (points).
        /// </summary>
        public void ArrangeShapesInGridAutoToSlideContentPoints(IEnumerable<PowerPointShape> shapes, double marginPoints,
            double gutterXPoints = 0d, double gutterYPoints = 0d, bool resizeToCell = true,
            PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            ArrangeShapesInGridAutoToSlideContent(shapes,
                PowerPointUnits.FromPoints(marginPoints),
                PowerPointUnits.FromPoints(gutterXPoints),
                PowerPointUnits.FromPoints(gutterYPoints),
                resizeToCell, flow);
        }

        private static PowerPointLayoutBox[,] BuildGrid(PowerPointLayoutBox bounds, int columns, int rows,
            long gutterX, long gutterY) {
            PowerPointLayoutBox[,] grid = new PowerPointLayoutBox[rows, columns];
            PowerPointLayoutBox[] rowBoxes = bounds.SplitRows(rows, gutterY);
            for (int r = 0; r < rows; r++) {
                PowerPointLayoutBox[] colBoxes = rowBoxes[r].SplitColumns(columns, gutterX);
                for (int c = 0; c < columns; c++) {
                    grid[r, c] = colBoxes[c];
                }
            }
            return grid;
        }

        private static int GetAutoGridColumns(int count, PowerPointLayoutBox bounds, long gutterX, long gutterY,
            int? minColumns, int? maxColumns, double? targetCellAspect) {
            if (count <= 1) {
                return 1;
            }

            int min = Math.Max(1, minColumns ?? 1);
            int max = Math.Min(count, maxColumns ?? count);

            if (min > max) {
                throw new ArgumentOutOfRangeException(nameof(minColumns), "MinColumns cannot be greater than MaxColumns.");
            }

            double targetGridRatio = bounds.Height == 0 ? count : bounds.Width / (double)bounds.Height;
            int bestColumns = min;
            double bestScore = double.MaxValue;
            int bestEmpty = int.MaxValue;
            double bestCellDelta = double.MaxValue;

            for (int columns = min; columns <= max; columns++) {
                int rows = (int)Math.Ceiling(count / (double)columns);
                double availableWidth = bounds.Width - gutterX * (columns - 1);
                double availableHeight = bounds.Height - gutterY * (rows - 1);
                if (availableWidth <= 0 || availableHeight <= 0) {
                    continue;
                }

                double cellWidth = availableWidth / columns;
                double cellHeight = availableHeight / rows;
                if (cellWidth <= 0 || cellHeight <= 0) {
                    continue;
                }

                double score;
                if (targetCellAspect.HasValue) {
                    score = Math.Abs((cellWidth / cellHeight) - targetCellAspect.Value);
                } else {
                    score = Math.Abs((columns / (double)rows) - targetGridRatio);
                }

                int empty = (columns * rows) - count;
                double cellDelta = Math.Abs((cellWidth / cellHeight) - 1d);

                if (score < bestScore ||
                    (Math.Abs(score - bestScore) < 0.0001 && (empty < bestEmpty ||
                                                             (empty == bestEmpty && cellDelta < bestCellDelta)))) {
                    bestScore = score;
                    bestEmpty = empty;
                    bestCellDelta = cellDelta;
                    bestColumns = columns;
                }
            }

            return bestColumns;
        }
    }
}
