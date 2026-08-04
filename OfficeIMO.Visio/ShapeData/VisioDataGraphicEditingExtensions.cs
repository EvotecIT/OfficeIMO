using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OfficeIMO.Visio {
    /// <summary>Typed snapshot of generated data-graphic shapes for one target.</summary>
    public sealed class VisioDataGraphicInstance {
        internal VisioDataGraphicInstance(VisioShape target,
            IReadOnlyList<VisioShape> shapes) {
            Target = target;
            Shapes = new ReadOnlyCollection<VisioShape>(new List<VisioShape>(shapes));
        }

        /// <summary>Shape whose Shape Data is visualized.</summary>
        public VisioShape Target { get; }

        /// <summary>Editable visible shapes that form the data graphic.</summary>
        public IReadOnlyList<VisioShape> Shapes { get; }

        /// <summary>Distinct Shape Data field names represented by the instance.</summary>
        public IReadOnlyList<string> FieldNames => Shapes
            .Select(shape => shape.GetUserCellValue(VisioSemanticUserCells.DataGraphicField))
            .Where(value => !string.IsNullOrWhiteSpace(value)).Cast<string>()
            .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    /// <summary>Typed data-graphic inspection, refresh, removal, and legend workflows.</summary>
    public static class VisioDataGraphicEditingExtensions {
        /// <summary>Returns generated data graphics currently attached to a target shape.</summary>
        public static VisioDataGraphicInstance GetDataGraphic(this VisioPage page,
            VisioShape target) {
            ValidateTarget(page, target);
            List<VisioShape> shapes = page.Shapes.Where(shape => string.Equals(
                    shape.GetUserCellValue(VisioSemanticUserCells.DataGraphicTargetId),
                    target.Id, StringComparison.OrdinalIgnoreCase))
                .ToList();
            return new VisioDataGraphicInstance(target, shapes);
        }

        /// <summary>Removes generated data-graphic shapes for one target.</summary>
        public static int RemoveDataGraphics(this VisioPage page,
            VisioShape target) {
            VisioDataGraphicInstance instance = page.GetDataGraphic(target);
            foreach (VisioShape shape in instance.Shapes) page.Shapes.Remove(shape);
            return instance.Shapes.Count;
        }

        /// <summary>Rebuilds one target's data graphics from current Shape Data values.</summary>
        public static IReadOnlyList<VisioShape> RefreshDataGraphics(
            this VisioPage page, VisioShape target,
            VisioDataGraphic definition) {
            ValidateTarget(page, target);
            if (definition == null) throw new ArgumentNullException(nameof(definition));
            VisioDataGraphicExtensions.ValidateDataGraphic(definition);
            List<(VisioShape Shape, int Index)> previous = page.GetDataGraphic(target)
                .Shapes.Select(shape => (shape, page.Shapes.IndexOf(shape)))
                .ToList();
            var before = new HashSet<VisioShape>(page.Shapes);
            foreach ((VisioShape shape, _) in previous) page.Shapes.Remove(shape);
            try {
                return page.AddDataGraphics(target, definition);
            } catch {
                for (int index = page.Shapes.Count - 1; index >= 0; index--) {
                    if (!before.Contains(page.Shapes[index])) page.Shapes.RemoveAt(index);
                }
                foreach ((VisioShape shape, int index) in previous
                             .OrderBy(item => item.Index)) {
                    page.Shapes.Insert(Math.Min(index, page.Shapes.Count), shape);
                }
                throw;
            }
        }

        /// <summary>Adds a compact typed legend for a data-graphic definition.</summary>
        public static VisioDataGraphicLegend AddDataGraphicLegend(
            this VisioPage page, string id, string title,
            VisioDataGraphic definition, double pinX, double pinY,
            double width = 2.4D) {
            if (page == null) throw new ArgumentNullException(nameof(page));
            if (definition == null) throw new ArgumentNullException(nameof(definition));
            if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Legend id cannot be empty.", nameof(id));
            if (string.IsNullOrWhiteSpace(title)) throw new ArgumentException("Legend title cannot be empty.", nameof(title));
            if (definition.Items.Count == 0) throw new ArgumentException("A legend requires at least one data-graphic item.", nameof(definition));
            if (width <= 0D || double.IsNaN(width) || double.IsInfinity(width)) throw new ArgumentOutOfRangeException(nameof(width));

            const double titleHeight = 0.35D;
            const double rowHeight = 0.32D;
            var shapes = new List<VisioShape>();
            double height = titleHeight + definition.Items.Count * rowHeight + 0.16D;
            var frame = new VisioShape(id, pinX, pinY, width, height, string.Empty) {
                Name = "Data Graphic Legend",
                NameU = "OfficeIMO Data Graphic Legend",
                FillColor = OfficeIMO.Drawing.OfficeColor.FromRgb(248, 250, 252),
                LineColor = OfficeIMO.Drawing.OfficeColor.FromRgb(148, 163, 184),
                LineWeight = 0.01D
            };
            frame.SetUserCell(VisioSemanticUserCells.Kind,
                VisioSemanticUserCells.DataGraphicLegendKind, "STR");
            page.Shapes.Add(frame);
            shapes.Add(frame);

            var heading = new VisioShape(id + "-title", pinX,
                pinY + height / 2D - titleHeight / 2D - 0.05D,
                width - 0.16D, titleHeight, title) {
                Name = "Data Graphic Legend Title",
                NameU = "OfficeIMO Data Graphic Legend Title",
                FillPattern = 0,
                LinePattern = 0
            };
            heading.SetUserCell(VisioSemanticUserCells.Kind,
                VisioSemanticUserCells.DataGraphicLegendKind, "STR");
            heading.SetUserCell(VisioSemanticUserCells.DataGraphicRole,
                VisioSemanticUserCells.DataGraphicLegendTitleRole, "STR");
            page.Shapes.Add(heading);
            shapes.Add(heading);

            for (int index = 0; index < definition.Items.Count; index++) {
                VisioDataGraphicItem item = definition.Items[index];
                double y = heading.PinY - titleHeight / 2D - rowHeight / 2D -
                           index * rowHeight;
                string legendText = string.IsNullOrWhiteSpace(item.Label)
                    ? item.FieldName
                    : item.Label!;
                var row = new VisioShape(id + "-item-" + index, pinX, y,
                    width - 0.2D, rowHeight - 0.04D,
                    legendText) {
                    Name = "Data Graphic Legend Item",
                    NameU = "OfficeIMO Data Graphic Legend Item",
                    FillPattern = 0,
                    LinePattern = 0
                };
                row.SetUserCell(VisioSemanticUserCells.Kind,
                    VisioSemanticUserCells.DataGraphicLegendKind, "STR");
                row.SetUserCell(VisioSemanticUserCells.DataGraphicRole,
                    VisioSemanticUserCells.DataGraphicLegendItemRole, "STR");
                row.SetUserCell(VisioSemanticUserCells.DataGraphicField,
                    item.FieldName, "STR");
                page.Shapes.Add(row);
                shapes.Add(row);
            }

            return new VisioDataGraphicLegend(frame, shapes);
        }

        private static void ValidateTarget(VisioPage page, VisioShape target) {
            if (page == null) throw new ArgumentNullException(nameof(page));
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (!page.AllShapes().Contains(target)) throw new InvalidOperationException("The data graphic target must belong to the page.");
        }
    }

    /// <summary>Editable generated legend for a data-graphic definition.</summary>
    public sealed class VisioDataGraphicLegend {
        internal VisioDataGraphicLegend(VisioShape frame,
            IReadOnlyList<VisioShape> shapes) {
            Frame = frame;
            Shapes = new ReadOnlyCollection<VisioShape>(new List<VisioShape>(shapes));
        }

        /// <summary>Legend frame.</summary>
        public VisioShape Frame { get; }

        /// <summary>Frame, title, and item shapes.</summary>
        public IReadOnlyList<VisioShape> Shapes { get; }
    }
}
