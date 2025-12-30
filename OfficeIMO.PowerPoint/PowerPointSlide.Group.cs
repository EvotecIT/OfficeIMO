using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Groups shapes into a single group shape.
        /// </summary>
        public PowerPointGroupShape GroupShapes(IEnumerable<PowerPointShape> shapes, string? name = null) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count < 2) {
                throw new InvalidOperationException("At least two shapes are required to create a group.");
            }

            List<PowerPointShape> ordered = list
                .OrderBy(shape => EnsureShapeOnSlide(shape))
                .ToList();

            OpenXmlElement? parent = ordered[0].Element.Parent;
            if (parent == null) {
                throw new InvalidOperationException("Shape is not attached to a slide.");
            }

            if (ordered.Any(shape => !ReferenceEquals(shape.Element.Parent, parent))) {
                throw new InvalidOperationException("All shapes must share the same parent.");
            }

            PowerPointLayoutBox bounds = GetSelectionBounds(ordered);
            string baseName = string.IsNullOrWhiteSpace(name) ? "Group" : name!;
            string groupName = GenerateUniqueName(baseName);

            GroupShape group = new GroupShape(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = _nextShapeId++, Name = groupName },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(
                    new A.TransformGroup(
                        new A.Offset { X = bounds.Left, Y = bounds.Top },
                        new A.Extents { Cx = bounds.Width, Cy = bounds.Height },
                        new A.ChildOffset { X = bounds.Left, Y = bounds.Top },
                        new A.ChildExtents { Cx = bounds.Width, Cy = bounds.Height }
                    )));

            OpenXmlElement firstElement = ordered[0].Element;
            parent.InsertBefore(group, firstElement);

            foreach (PowerPointShape shape in ordered) {
                shape.Element.Remove();
                group.Append(shape.Element);
            }

            int insertIndex = _shapes.IndexOf(ordered[0]);
            foreach (PowerPointShape shape in ordered) {
                _shapes.Remove(shape);
            }

            PowerPointGroupShape grouped = new PowerPointGroupShape(group);
            _shapes.Insert(insertIndex, grouped);
            return grouped;
        }

        /// <summary>
        ///     Ungroups a group shape back into its child shapes.
        /// </summary>
        public IReadOnlyList<PowerPointShape> UngroupShape(PowerPointGroupShape groupShape) {
            if (groupShape == null) {
                throw new ArgumentNullException(nameof(groupShape));
            }

            int index = EnsureShapeOnSlide(groupShape);
            GroupShape group = groupShape.GroupShape;
            OpenXmlElement? parent = group.Parent;
            if (parent == null) {
                throw new InvalidOperationException("Group is not attached to a slide.");
            }

            List<OpenXmlElement> children = group.ChildElements
                .Where(child => child is not NonVisualGroupShapeProperties && child is not GroupShapeProperties)
                .ToList();

            var results = new List<PowerPointShape>();
            foreach (OpenXmlElement child in children) {
                child.Remove();
                parent.InsertBefore(child, group);
                PowerPointShape? shape = CreateShapeFromElement(child);
                if (shape != null) {
                    results.Add(shape);
                }
            }

            group.Remove();

            _shapes.RemoveAt(index);
            _shapes.InsertRange(index, results);
            return results;
        }
    }
}
