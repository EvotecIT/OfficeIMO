using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Groups shapes into a single group shape.
        /// </summary>
        public PowerPointGroupShape GroupShapes(IEnumerable<PowerPointShape> shapes, string? name = null) {
            ArgumentNullException.ThrowIfNull(shapes);

            List<PowerPointShape> list = NormalizeShapes(shapes);
            if (list.Count < 2) {
                throw new InvalidOperationException("At least two shapes are required to create a group.");
            }

            List<PowerPointShape> ordered = list
                .OrderBy(shape => EnsureShapeOnSlide(shape))
                .ToList();

            PowerPointPartRef part = PartRef.ForMovingXmlNodes();
            OfficeXmlNode parent = ordered[0].XmlElement.Parent;
            if (parent.IsNone) {
                throw new InvalidOperationException("Shape is not attached to a slide.");
            }

            if (ordered.Any(shape => shape.XmlElement.Parent != parent)) {
                throw new InvalidOperationException("All shapes must share the same parent.");
            }

            PowerPointLayoutBox bounds = GetSelectionBounds(ordered);
            string baseName = string.IsNullOrWhiteSpace(name) ? "Group" : name!;
            string groupName = GenerateUniqueName(baseName);

            const int presentation = OfficeXmlNamespaces.PresentationML;
            const int drawing = OfficeXmlNamespaces.DrawingML;
            uint groupId = AllocateShapeId();
            // An existing slide may have needed the SDK when resolving its drawing IDs.
            parent = ordered[0].XmlElement.Parent;
            part = part.Current;
            OfficeXmlNode group = parent.CreateElement(presentation, "grpSp");
            OfficeXmlNode nonVisual = group.Append(group.CreateElement(presentation, "nvGrpSpPr"));
            OfficeXmlNode properties = nonVisual.Append(nonVisual.CreateElement(presentation, "cNvPr"));
            properties.SetUInt32Attribute("id", groupId);
            properties.SetAttribute("name", groupName);
            nonVisual.Append(nonVisual.CreateElement(presentation, "cNvGrpSpPr"));
            nonVisual.Append(nonVisual.CreateElement(presentation, "nvPr"));

            OfficeXmlNode groupProperties = group.Append(group.CreateElement(presentation, "grpSpPr"));
            OfficeXmlNode transform = groupProperties.Append(groupProperties.CreateElement(drawing, "xfrm"));
            WriteGroupPoint(transform, "off", "x", bounds.Left, "y", bounds.Top);
            WriteGroupPoint(transform, "ext", "cx", bounds.Width, "cy", bounds.Height);
            WriteGroupPoint(transform, "chOff", "x", bounds.Left, "y", bounds.Top);
            WriteGroupPoint(transform, "chExt", "cx", bounds.Width, "cy", bounds.Height);
            group = parent.InsertBefore(group, ordered[0].XmlElement);

            foreach (PowerPointShape shape in ordered) {
                OfficeXmlNode element = shape.XmlElement;
                element.Remove();
                group.Append(element);
            }

            int insertIndex = ShapeList.IndexOf(ordered[0]);
            foreach (PowerPointShape shape in ordered) {
                // The caller may still hold a child after it leaves the slide's top-level list.
                if (shape.IsPackageBacked) OwnerDocument?.RegisterPackageBinding(shape);
                ShapeList.Remove(shape);
            }

            var grouped = (PowerPointGroupShape)WrapSlideShape(group, part)!;
            part.MarkEdited();
            ShapeList.Insert(insertIndex, grouped);
            return grouped;
        }

        /// <summary>
        ///     Ungroups a group shape back into its child shapes.
        /// </summary>
        public IReadOnlyList<PowerPointShape> UngroupShape(PowerPointGroupShape groupShape) {
            ArgumentNullException.ThrowIfNull(groupShape);

            int index = EnsureShapeOnSlide(groupShape);
            PowerPointPartRef part = PartRef.ForMovingXmlNodes();
            OfficeXmlNode group = groupShape.XmlElement;
            OfficeXmlNode parent = group.Parent;
            if (parent.IsNone) {
                throw new InvalidOperationException("Group is not attached to a slide.");
            }

            const int presentation = OfficeXmlNamespaces.PresentationML;
            var results = new List<PowerPointShape>();
            foreach (OfficeXmlNode child in group.ChildNodes()) {
                if (child.Is(presentation, "nvGrpSpPr") || child.Is(presentation, "grpSpPr")) continue;
                child.Remove();
                OfficeXmlNode moved = parent.InsertBefore(child, group);
                if (WrapSlideShape(moved, part) is { } shape) results.Add(shape);
            }

            group.Remove();
            part.MarkEdited();

            ShapeList.RemoveAt(index);
            ShapeList.InsertRange(index, results);
            return results;
        }
        private static void WriteGroupPoint(OfficeXmlNode transform, string name,
            string horizontalAttribute, OfficeLength horizontal, string verticalAttribute, OfficeLength vertical) {
            OfficeXmlNode point = transform.Append(transform.CreateElement(OfficeXmlNamespaces.DrawingML, name));
            point.SetInt64Attribute(horizontalAttribute, horizontal.ToEmus());
            point.SetInt64Attribute(verticalAttribute, vertical.ToEmus());
        }
    }
}
