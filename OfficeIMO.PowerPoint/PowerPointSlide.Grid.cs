using System;
using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Arranges shapes into a grid within their selection bounds.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="columns">Number of grid columns, greater than zero.</param>
        /// <param name="rows">Number of grid rows, greater than zero.</param>
        /// <param name="gutterX">Nonnegative horizontal gap between cells. The default is zero.</param>
        /// <param name="gutterY">Nonnegative vertical gap between cells. The default is zero.</param>
        /// <param name="resizeToCell">Whether each shape is resized to its grid cell.</param>
        /// <param name="flow">Order used to fill grid cells.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void ArrangeShapesInGrid(IEnumerable<PowerPointShape> shapes, int columns, int rows,
            OfficeLength gutterX = default, OfficeLength gutterY = default, bool resizeToCell = true,
            PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            ArrangeShapesInGrid(shapes, GetSelectionBounds(shapes), columns, rows, gutterX, gutterY, resizeToCell, flow);
        }

        /// <summary>
        ///     Arranges shapes into a grid within a custom bounding box.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="bounds">Target region; positions may be signed and dimensions must be nonnegative.</param>
        /// <param name="columns">Number of grid columns, greater than zero.</param>
        /// <param name="rows">Number of grid rows, greater than zero.</param>
        /// <param name="gutterX">Nonnegative horizontal gap between cells. The default is zero.</param>
        /// <param name="gutterY">Nonnegative vertical gap between cells. The default is zero.</param>
        /// <param name="resizeToCell">Whether each shape is resized to its grid cell.</param>
        /// <param name="flow">Order used to fill grid cells.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void ArrangeShapesInGrid(IEnumerable<PowerPointShape> shapes, PowerPointLayoutBox bounds, int columns,
            int rows, OfficeLength gutterX = default, OfficeLength gutterY = default, bool resizeToCell = true,
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
            if (gutterX < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(gutterX));
            }
            if (gutterY < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(gutterY));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count == 0) {
                return;
            }

            ValidateLayoutBounds(bounds);
            ValidateGridFlow(flow);
            PowerPointLayoutBox[,] grid = BuildGrid(bounds, columns, rows, gutterX.ToEmus(), gutterY.ToEmus());
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
                    shape.Left= OfficeIMO.OfficeLength.Emus(cell.Left.ToEmus());
                    shape.Top= OfficeIMO.OfficeLength.Emus(cell.Top.ToEmus());
                }
            }
        }

        /// <summary>
        ///     Arranges shapes into a grid that spans the slide bounds.        
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="columns">Number of grid columns, greater than zero.</param>
        /// <param name="rows">Number of grid rows, greater than zero.</param>
        /// <param name="gutterX">Nonnegative horizontal gap between cells. The default is zero.</param>
        /// <param name="gutterY">Nonnegative vertical gap between cells. The default is zero.</param>
        /// <param name="resizeToCell">Whether each shape is resized to its grid cell.</param>
        /// <param name="flow">Order used to fill grid cells.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void ArrangeShapesInGridToSlide(IEnumerable<PowerPointShape> shapes, int columns, int rows,
            OfficeLength gutterX = default, OfficeLength gutterY = default, bool resizeToCell = true,
            PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {  
            ArrangeShapesInGrid(shapes, GetSlideBounds(), columns, rows, gutterX, gutterY, resizeToCell, flow);
        }

        /// <summary>
        ///     Arranges shapes into a grid that spans the slide content bounds using a margin.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="columns">Number of grid columns, greater than zero.</param>
        /// <param name="rows">Number of grid rows, greater than zero.</param>
        /// <param name="margin">Nonnegative inset on every side of the slide.</param>
        /// <param name="gutterX">Nonnegative horizontal gap between cells. The default is zero.</param>
        /// <param name="gutterY">Nonnegative vertical gap between cells. The default is zero.</param>
        /// <param name="resizeToCell">Whether each shape is resized to its grid cell.</param>
        /// <param name="flow">Order used to fill grid cells.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void ArrangeShapesInGridToSlideContent(IEnumerable<PowerPointShape> shapes, int columns, int rows,
            OfficeLength margin, OfficeLength gutterX = default, OfficeLength gutterY = default, bool resizeToCell = true,
            PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            ArrangeShapesInGrid(shapes, GetSlideContentBounds(margin), columns, rows, gutterX, gutterY,
                resizeToCell, flow);
        }

        /// <summary>
        ///     Arranges shapes into an auto-sized grid within their selection bounds.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="gutterX">Nonnegative horizontal gap between cells. The default is zero.</param>
        /// <param name="gutterY">Nonnegative vertical gap between cells. The default is zero.</param>
        /// <param name="resizeToCell">Whether each shape is resized to its grid cell.</param>
        /// <param name="flow">Order used to fill grid cells.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void ArrangeShapesInGridAuto(IEnumerable<PowerPointShape> shapes, OfficeLength gutterX = default, OfficeLength gutterY = default,
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
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="bounds">Target region; positions may be signed and dimensions must be nonnegative.</param>
        /// <param name="gutterX">Nonnegative horizontal gap between cells. The default is zero.</param>
        /// <param name="gutterY">Nonnegative vertical gap between cells. The default is zero.</param>
        /// <param name="resizeToCell">Whether each shape is resized to its grid cell.</param>
        /// <param name="flow">Order used to fill grid cells.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void ArrangeShapesInGridAuto(IEnumerable<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            OfficeLength gutterX = default, OfficeLength gutterY = default, bool resizeToCell = true,
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
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="bounds">Target region; positions may be signed and dimensions must be nonnegative.</param>
        /// <param name="options">Layout settings, or null to use the default settings.</param>
        public void ArrangeShapesInGridAuto(IEnumerable<PowerPointShape> shapes, PowerPointLayoutBox bounds,
            PowerPointShapeGridOptions? options) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count == 0) {
                return;
            }

            PowerPointShapeGridOptions resolvedOptions = options ?? new PowerPointShapeGridOptions();
            ValidateLayoutBounds(bounds);
            if (resolvedOptions.GutterX < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(options), "Horizontal gutter cannot be negative.");
            if (resolvedOptions.GutterY < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(options), "Vertical gutter cannot be negative.");
            ValidateGridFlow(resolvedOptions.Flow);
            int columns = GetAutoGridColumns(list.Count, bounds, resolvedOptions.GutterX.ToEmus(), resolvedOptions.GutterY.ToEmus(),
                resolvedOptions.MinColumns, resolvedOptions.MaxColumns, resolvedOptions.TargetCellAspect);
            int rows = (int)Math.Ceiling(list.Count / (double)columns);

            ArrangeShapesInGrid(list, bounds, columns, rows, resolvedOptions.GutterX, resolvedOptions.GutterY,
                resolvedOptions.ResizeToCell, resolvedOptions.Flow);
        }

        /// <summary>
        ///     Arranges shapes into an auto-sized grid that spans the slide bounds.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="gutterX">Nonnegative horizontal gap between cells. The default is zero.</param>
        /// <param name="gutterY">Nonnegative vertical gap between cells. The default is zero.</param>
        /// <param name="resizeToCell">Whether each shape is resized to its grid cell.</param>
        /// <param name="flow">Order used to fill grid cells.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void ArrangeShapesInGridAutoToSlide(IEnumerable<PowerPointShape> shapes, OfficeLength gutterX = default, OfficeLength gutterY = default,
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
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="options">Layout settings, or null to use the default settings.</param>
        public void ArrangeShapesInGridAutoToSlide(IEnumerable<PowerPointShape> shapes, PowerPointShapeGridOptions? options) {
            ArrangeShapesInGridAuto(shapes, GetSlideBounds(), options);
        }

        /// <summary>
        ///     Arranges shapes into an auto-sized grid that spans the slide content bounds using a margin.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="margin">Nonnegative inset on every side of the slide.</param>
        /// <param name="gutterX">Nonnegative horizontal gap between cells. The default is zero.</param>
        /// <param name="gutterY">Nonnegative vertical gap between cells. The default is zero.</param>
        /// <param name="resizeToCell">Whether each shape is resized to its grid cell.</param>
        /// <param name="flow">Order used to fill grid cells.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void ArrangeShapesInGridAutoToSlideContent(IEnumerable<PowerPointShape> shapes, OfficeLength margin,
            OfficeLength gutterX = default, OfficeLength gutterY = default, bool resizeToCell = true,
            PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            ArrangeShapesInGridAuto(shapes, GetSlideContentBounds(margin),
                new PowerPointShapeGridOptions {
                    GutterX = gutterX,
                    GutterY = gutterY,
                    ResizeToCell = resizeToCell,
                    Flow = flow
                });
        }

        /// <summary>
        ///     Arranges shapes into an auto-sized grid that spans the slide content bounds using a margin and options.
        /// </summary>
        /// <param name="shapes">Shapes to arrange; null entries are ignored.</param>
        /// <param name="margin">Nonnegative inset on every side of the slide.</param>
        /// <param name="options">Layout settings, or null to use the default settings.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void ArrangeShapesInGridAutoToSlideContent(IEnumerable<PowerPointShape> shapes, OfficeLength margin,
            PowerPointShapeGridOptions? options) {
            ArrangeShapesInGridAuto(shapes, GetSlideContentBounds(margin), options);
        }

        private static void ValidateGridFlow(PowerPointShapeGridFlow flow) {
            if (flow != PowerPointShapeGridFlow.RowMajor && flow != PowerPointShapeGridFlow.ColumnMajor) {
                throw new ArgumentOutOfRangeException(nameof(flow));
            }
        }

        private static PowerPointLayoutBox[,] BuildGrid(PowerPointLayoutBox bounds, int columns, int rows,
            long gutterX, long gutterY) {
            return bounds.SplitGrid(rows,columns,OfficeIMO.OfficeLength.Emus(gutterY),OfficeIMO.OfficeLength.Emus(gutterX));
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

            double targetGridRatio = bounds.Height.ToEmus() == 0 ? count : bounds.Width.ToEmus() / (double)bounds.Height.ToEmus();
            int bestColumns = min;
            double bestScore = double.MaxValue;
            int bestEmpty = int.MaxValue;
            double bestCellDelta = double.MaxValue;

            for (int columns = min; columns <= max; columns++) {
                int rows = (int)Math.Ceiling(count / (double)columns);
                double availableWidth = bounds.Width.ToEmus() - (double)gutterX * (columns - 1);
                double availableHeight = bounds.Height.ToEmus() - (double)gutterY * (rows - 1);
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
