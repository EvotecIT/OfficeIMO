using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Returns the child shapes of a group shape.
        /// </summary>
        public IReadOnlyList<PowerPointShape> GetGroupChildren(PowerPointGroupShape groupShape) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }

            GroupShape group = groupShape.GroupShape;
            var children = new List<PowerPointShape>();
            foreach (var child in group.ChildElements) {
                if (child is NonVisualGroupShapeProperties || child is GroupShapeProperties) {
                    continue;
                }

                PowerPointShape? shape = CreateShapeFromElement(child);
                if (shape != null) {
                    children.Add(shape);
                }
            }

            return children;
        }

        /// <summary>
        ///     Returns the child textboxes of a group shape.
        /// </summary>
        public IReadOnlyList<PowerPointTextBox> GetGroupTextBoxes(PowerPointGroupShape groupShape) {
            return GetGroupChildren(groupShape).OfType<PowerPointTextBox>().ToList();
        }

        /// <summary>
        ///     Returns the child pictures of a group shape.
        /// </summary>
        public IReadOnlyList<PowerPointPicture> GetGroupPictures(PowerPointGroupShape groupShape) {
            return GetGroupChildren(groupShape).OfType<PowerPointPicture>().ToList();
        }

        /// <summary>
        ///     Returns the child tables of a group shape.
        /// </summary>
        public IReadOnlyList<PowerPointTable> GetGroupTables(PowerPointGroupShape groupShape) {
            return GetGroupChildren(groupShape).OfType<PowerPointTable>().ToList();
        }

        /// <summary>
        ///     Returns the child charts of a group shape.
        /// </summary>
        public IReadOnlyList<PowerPointChart> GetGroupCharts(PowerPointGroupShape groupShape) {
            return GetGroupChildren(groupShape).OfType<PowerPointChart>().ToList();
        }

        /// <summary>
        ///     Gets the layout bounds used by a group for its children.
        /// </summary>
        public PowerPointLayoutBox GetGroupChildBounds(PowerPointGroupShape groupShape) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }

            IReadOnlyList<PowerPointShape> children = GetGroupChildren(groupShape);
            return GetGroupChildBounds(groupShape, children);
        }

        /// <summary>
        ///     Gets the layout bounds used by a group for its children in centimeters.
        /// </summary>
        public PowerPointLayoutBox GetGroupChildBoundsCm(PowerPointGroupShape groupShape) {
            PowerPointLayoutBox bounds = GetGroupChildBounds(groupShape);
            return PowerPointLayoutBox.FromCentimeters(bounds.LeftCm, bounds.TopCm, bounds.WidthCm, bounds.HeightCm);
        }

        /// <summary>
        ///     Gets the layout bounds used by a group for its children in inches.
        /// </summary>
        public PowerPointLayoutBox GetGroupChildBoundsInches(PowerPointGroupShape groupShape) {
            PowerPointLayoutBox bounds = GetGroupChildBounds(groupShape);
            return PowerPointLayoutBox.FromInches(bounds.LeftInches, bounds.TopInches, bounds.WidthInches, bounds.HeightInches);
        }

        /// <summary>
        ///     Gets the layout bounds used by a group for its children in points.
        /// </summary>
        public PowerPointLayoutBox GetGroupChildBoundsPoints(PowerPointGroupShape groupShape) {
            PowerPointLayoutBox bounds = GetGroupChildBounds(groupShape);
            return PowerPointLayoutBox.FromPoints(bounds.LeftPoints, bounds.TopPoints, bounds.WidthPoints, bounds.HeightPoints);
        }

        /// <summary>
        ///     Aligns child shapes within the group's bounds.
        /// </summary>
        public void AlignGroupChildren(PowerPointGroupShape groupShape, PowerPointShapeAlignment alignment) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }

            List<PowerPointShape> children = GetGroupChildren(groupShape).ToList();
            if (children.Count == 0) {
                return;
            }

            AlignShapes(children, alignment, GetGroupChildBounds(groupShape, children));
        }

        /// <summary>
        ///     Distributes child shapes within the group's bounds.
        /// </summary>
        public void DistributeGroupChildren(PowerPointGroupShape groupShape, PowerPointShapeDistribution distribution) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }

            List<PowerPointShape> children = GetGroupChildren(groupShape).ToList();
            if (children.Count < 2) {
                return;
            }

            DistributeShapes(children, distribution, GetGroupChildBounds(groupShape, children));
        }

        /// <summary>
        ///     Distributes child shapes within the group's bounds and aligns on the cross axis.
        /// </summary>
        public void DistributeGroupChildren(PowerPointGroupShape groupShape, PowerPointShapeDistribution distribution,
            PowerPointShapeAlignment crossAxisAlignment) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }

            List<PowerPointShape> children = GetGroupChildren(groupShape).ToList();
            if (children.Count < 2) {
                return;
            }

            DistributeShapes(children, distribution, GetGroupChildBounds(groupShape, children), crossAxisAlignment);
        }

        /// <summary>
        ///     Distributes child shapes with fixed spacing within the group's bounds.
        /// </summary>
        public void DistributeGroupChildrenWithSpacing(PowerPointGroupShape groupShape, PowerPointShapeDistribution distribution,
            long spacingEmus, bool center = false) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }

            List<PowerPointShape> children = GetGroupChildren(groupShape).ToList();
            if (children.Count < 2) {
                return;
            }

            DistributeShapesWithSpacing(children, distribution, GetGroupChildBounds(groupShape, children),
                spacingEmus, center);
        }

        /// <summary>
        ///     Distributes child shapes with fixed spacing within the group's bounds using options.
        /// </summary>
        public void DistributeGroupChildrenWithSpacing(PowerPointGroupShape groupShape, PowerPointShapeDistribution distribution,
            PowerPointShapeSpacingOptions options) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }

            List<PowerPointShape> children = GetGroupChildren(groupShape).ToList();
            if (children.Count < 2) {
                return;
            }

            DistributeShapesWithSpacing(children, distribution, GetGroupChildBounds(groupShape, children), options);
        }

        /// <summary>
        ///     Stacks child shapes within the group's bounds.
        /// </summary>
        public void StackGroupChildren(PowerPointGroupShape groupShape, PowerPointShapeStackDirection direction, long spacingEmus) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }

            List<PowerPointShape> children = GetGroupChildren(groupShape).ToList();
            if (children.Count == 0) {
                return;
            }

            StackShapes(children, direction, GetGroupChildBounds(groupShape, children), spacingEmus);
        }

        /// <summary>
        ///     Stacks child shapes within the group's bounds using options.
        /// </summary>
        public void StackGroupChildren(PowerPointGroupShape groupShape, PowerPointShapeStackDirection direction,
            PowerPointShapeStackOptions options) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }

            List<PowerPointShape> children = GetGroupChildren(groupShape).ToList();
            if (children.Count == 0) {
                return;
            }

            StackShapes(children, direction, GetGroupChildBounds(groupShape, children), options);
        }

        /// <summary>
        ///     Arranges child shapes into a grid within the group's bounds.
        /// </summary>
        public void ArrangeGroupChildrenInGrid(PowerPointGroupShape groupShape, int columns, int rows,
            long gutterX = 0L, long gutterY = 0L, bool resizeToCell = true,
            PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }

            List<PowerPointShape> children = GetGroupChildren(groupShape).ToList();
            if (children.Count == 0) {
                return;
            }

            ArrangeShapesInGrid(children, GetGroupChildBounds(groupShape, children), columns, rows,
                gutterX, gutterY, resizeToCell, flow);
        }

        /// <summary>
        ///     Arranges child shapes into an auto-sized grid within the group's bounds.
        /// </summary>
        public void ArrangeGroupChildrenInGridAuto(PowerPointGroupShape groupShape, long gutterX = 0L, long gutterY = 0L,
            bool resizeToCell = true, PowerPointShapeGridFlow flow = PowerPointShapeGridFlow.RowMajor) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }

            List<PowerPointShape> children = GetGroupChildren(groupShape).ToList();
            if (children.Count == 0) {
                return;
            }

            ArrangeShapesInGridAuto(children, GetGroupChildBounds(groupShape, children),
                gutterX, gutterY, resizeToCell, flow);
        }

        /// <summary>
        ///     Arranges child shapes into an auto-sized grid within the group's bounds using options.
        /// </summary>
        public void ArrangeGroupChildrenInGridAuto(PowerPointGroupShape groupShape, PowerPointShapeGridOptions options) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }

            List<PowerPointShape> children = GetGroupChildren(groupShape).ToList();
            if (children.Count == 0) {
                return;
            }

            ArrangeShapesInGridAuto(children, GetGroupChildBounds(groupShape, children), options);
        }

        private static PowerPointLayoutBox GetGroupChildBounds(PowerPointGroupShape groupShape,
            IReadOnlyList<PowerPointShape> children) {
            A.TransformGroup? transform = groupShape.GroupShape.GroupShapeProperties?.TransformGroup;
            long? x = transform?.ChildOffset?.X?.Value;
            long? y = transform?.ChildOffset?.Y?.Value;
            long? cx = transform?.ChildExtents?.Cx?.Value;
            long? cy = transform?.ChildExtents?.Cy?.Value;

            if (x.HasValue && y.HasValue && cx.HasValue && cy.HasValue) {
                return new PowerPointLayoutBox(x.Value, y.Value, cx.Value, cy.Value);
            }

            if (children.Count == 0) {
                return new PowerPointLayoutBox(0, 0, 0, 0);
            }

            return GetSelectionBounds(children);
        }
    }
}
