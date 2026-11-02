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
        public IReadOnlyList<PowerPointShape> GetGroupChildren(PowerPointGroupShape groupShape) =>
            GetGroupChildren(groupShape, int.MaxValue);

        internal IReadOnlyList<PowerPointShape> GetGroupChildren(PowerPointGroupShape groupShape,
            int maximumChildren) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }
            if (maximumChildren < 0) throw new ArgumentOutOfRangeException(nameof(maximumChildren));

            GroupShape group = groupShape.GroupShape;
            var children = new List<PowerPointShape>();
            foreach (var child in group.ChildElements) {
                if (child is NonVisualGroupShapeProperties || child is GroupShapeProperties) {
                    continue;
                }

                PowerPointShape? shape = CreateShapeFromElement(child, groupShape.OwnerPart ?? _slidePart);
                if (shape != null) {
                    if (children.Count >= maximumChildren) {
                        throw new InvalidOperationException("The shape count exceeds the configured inspection limit.");
                    }
                    children.Add(shape.AttachTo(this, groupShape.OwnerPart));
                }
            }

            return children;
        }

        /// <summary>
        ///     Enumerates a shape tree in drawing order, including nested group children while preserving
        ///     inherited hidden state.
        /// </summary>
        public IEnumerable<PowerPointShape> EnumerateShapesDeep(IEnumerable<PowerPointShape> shapes,
            bool includeHidden = false) {
            return EnumerateShapesDeep(shapes, includeHidden, 10000, 128);
        }

        internal IEnumerable<PowerPointShape> EnumerateShapesDeep(IEnumerable<PowerPointShape> shapes,
            bool includeHidden, int maximumShapeCount, int maximumGroupDepth) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }

            var stack = new Stack<(IEnumerator<PowerPointShape> Enumerator, int Depth, bool CountOnMove)>();
            stack.Push((shapes.GetEnumerator(), 0, true));
            int scheduledShapeCount = 0;
            int shapeCount = 0;
            try {
                while (stack.Count > 0) {
                    (IEnumerator<PowerPointShape> enumerator, int depth, bool countOnMove) = stack.Peek();
                    if (!enumerator.MoveNext()) {
                        stack.Pop().Enumerator.Dispose();
                        continue;
                    }
                    if (countOnMove && ++scheduledShapeCount > maximumShapeCount) {
                        throw new InvalidOperationException("The shape count exceeds the configured inspection limit.");
                    }
                    PowerPointShape shape = enumerator.Current;
                    if (!includeHidden && shape.Hidden) continue;
                    if (depth > maximumGroupDepth) {
                        throw new InvalidOperationException("The grouped-shape nesting depth exceeds the configured limit.");
                    }
                    if (++shapeCount > maximumShapeCount) {
                        throw new InvalidOperationException("The shape count exceeds the configured inspection limit.");
                    }
                    yield return shape;
                    if (shape is not PowerPointGroupShape groupShape) continue;
                    int remainingShapeCount = maximumShapeCount - scheduledShapeCount;
                    IReadOnlyList<PowerPointShape> children = GetGroupChildren(groupShape, remainingShapeCount);
                    scheduledShapeCount += children.Count;
                    stack.Push((children.GetEnumerator(), depth + 1, false));
                }
            } finally {
                while (stack.Count > 0) stack.Pop().Enumerator.Dispose();
            }
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
        ///     Gets the layout bounds used by a group for its children as OfficeLength values.
        /// </summary>
        /// <param name="groupShape">Group whose native child coordinate space is inspected.</param>
        /// <returns>The native child region, or the selection bounds when no complete child transform exists.</returns>
        public PowerPointLayoutBox GetGroupChildBounds(PowerPointGroupShape groupShape) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }

            IReadOnlyList<PowerPointShape> children = GetGroupChildren(groupShape);
            return GetGroupChildBounds(groupShape, children);
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
        /// <param name="groupShape">Group whose children are arranged in their native child coordinate space.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <param name="center">Whether to center the arranged shapes within the target bounds.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void DistributeGroupChildrenWithSpacing(PowerPointGroupShape groupShape, PowerPointShapeDistribution distribution,
            OfficeLength spacing, bool center = false) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }

            List<PowerPointShape> children = GetGroupChildren(groupShape).ToList();
            if (children.Count < 2) {
                return;
            }

            DistributeShapesWithSpacing(children, distribution, GetGroupChildBounds(groupShape, children),
                spacing, center);
        }

        /// <summary>
        ///     Distributes child shapes with fixed spacing within the group's bounds using options.
        /// </summary>
        /// <param name="groupShape">Group whose children are arranged in their native child coordinate space.</param>
        /// <param name="distribution">Axis along which shapes are distributed.</param>
        /// <param name="options">Layout settings, or null to use the default settings.</param>
        public void DistributeGroupChildrenWithSpacing(PowerPointGroupShape groupShape, PowerPointShapeDistribution distribution,
            PowerPointShapeSpacingOptions? options) {
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
        /// <param name="groupShape">Group whose children are arranged in their native child coordinate space.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="spacing">Nonnegative distance between adjacent shapes.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void StackGroupChildren(PowerPointGroupShape groupShape, PowerPointShapeStackDirection direction, OfficeLength spacing) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }

            List<PowerPointShape> children = GetGroupChildren(groupShape).ToList();
            if (children.Count == 0) {
                return;
            }

            StackShapes(children, direction, GetGroupChildBounds(groupShape, children), spacing);
        }

        /// <summary>
        ///     Stacks child shapes within the group's bounds using options.
        /// </summary>
        /// <param name="groupShape">Group whose children are arranged in their native child coordinate space.</param>
        /// <param name="direction">Axis along which the stack grows.</param>
        /// <param name="options">Layout settings, or null to use the default settings.</param>
        public void StackGroupChildren(PowerPointGroupShape groupShape, PowerPointShapeStackDirection direction,
            PowerPointShapeStackOptions? options) {
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
        /// <param name="groupShape">Group whose children are arranged in their native child coordinate space.</param>
        /// <param name="columns">Number of grid columns, greater than zero.</param>
        /// <param name="rows">Number of grid rows, greater than zero.</param>
        /// <param name="gutterX">Nonnegative horizontal gap between cells. The default is zero.</param>
        /// <param name="gutterY">Nonnegative vertical gap between cells. The default is zero.</param>
        /// <param name="resizeToCell">Whether each shape is resized to its grid cell.</param>
        /// <param name="flow">Order used to fill grid cells.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void ArrangeGroupChildrenInGrid(PowerPointGroupShape groupShape, int columns, int rows,
            OfficeLength gutterX = default, OfficeLength gutterY = default, bool resizeToCell = true,
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
        /// <param name="groupShape">Group whose children are arranged in their native child coordinate space.</param>
        /// <param name="gutterX">Nonnegative horizontal gap between cells. The default is zero.</param>
        /// <param name="gutterY">Nonnegative vertical gap between cells. The default is zero.</param>
        /// <param name="resizeToCell">Whether each shape is resized to its grid cell.</param>
        /// <param name="flow">Order used to fill grid cells.</param>
        /// <remarks>Geometry uses whole native EMUs. OfficeLength factories specify the input unit.</remarks>
        public void ArrangeGroupChildrenInGridAuto(PowerPointGroupShape groupShape, OfficeLength gutterX = default, OfficeLength gutterY = default,
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
        /// <param name="groupShape">Group whose children are arranged in their native child coordinate space.</param>
        /// <param name="options">Layout settings, or null to use the default settings.</param>
        public void ArrangeGroupChildrenInGridAuto(PowerPointGroupShape groupShape, PowerPointShapeGridOptions? options) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }

            List<PowerPointShape> children = GetGroupChildren(groupShape).ToList();
            if (children.Count == 0) {
                return;
            }

            ArrangeShapesInGridAuto(children, GetGroupChildBounds(groupShape, children), options);
        }

        internal static PowerPointLayoutBox GetGroupChildBounds(PowerPointGroupShape groupShape,
            IReadOnlyList<PowerPointShape> children) {
            A.TransformGroup? transform = groupShape.GroupShape.GroupShapeProperties?.TransformGroup;
            long? x = transform?.ChildOffset?.X?.Value;
            long? y = transform?.ChildOffset?.Y?.Value;
            long? cx = transform?.ChildExtents?.Cx?.Value;
            long? cy = transform?.ChildExtents?.Cy?.Value;

            if (x.HasValue && y.HasValue && cx.HasValue && cy.HasValue) {
                return new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(x.Value),OfficeIMO.OfficeLength.Emus(y.Value),OfficeIMO.OfficeLength.Emus(cx.Value),OfficeIMO.OfficeLength.Emus(cy.Value));
            }

            if (children.Count == 0) {
                return new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0));
            }

            return GetSelectionBounds(children);
        }
    }
}
