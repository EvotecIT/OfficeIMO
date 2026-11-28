using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private PowerPointShape? TryDuplicatePackageShape(PowerPointShape shape, int index, OfficeLength offsetX, OfficeLength offsetY) {
            // Charts and OLE objects duplicate their own dependent parts.
            if (shape is PowerPointChart or PowerPointOleObject || !shape.TryEditPackage(out _, out _)) return null;
            PowerPointPartRef part = PartRef.ForMovingXmlNodes();
            EnsureShapeIdsResolved();
            part = part.Current;
            if (part.IsSdk) return null;

            OfficeXmlNode original = shape.XmlElement;
            OfficeXmlNode parent = RequireShapeParent(original);
            if (!HasBoundedCopyDepth(original)) return null;

            bool shifted = offsetX != OfficeLength.Zero || offsetY != OfficeLength.Zero;
            OfficeLength left = shifted ? shape.Left + offsetX : default;
            OfficeLength top = shifted ? shape.Top + offsetY : default;
            uint firstId = AllocateShapeIds(ShapeNodesToRenumber(original).Count());

            OfficeXmlNode clone = original.Clone();
            PowerPointShape duplicate = WrapSlideShape(clone, part)
                ?? throw new InvalidOperationException("Duplicated shape type is not supported.");
            int position = -1;
            foreach (OfficeXmlNode node in ShapeNodesToRenumber(clone)) {
                ++position;
                OfficeXmlNode drawing = PowerPointShape.NonVisualDrawingElement(node, create: false);
                string baseName = position == 0 ? GetDuplicateBaseName(shape) : GroupMemberBaseName(node, drawing);
                if (drawing.IsNone) drawing = PowerPointShape.NonVisualDrawingElement(node, create: true);
                if (drawing.IsNone) continue;
                drawing.SetUInt32Attribute("id", firstId + (uint)position);
                drawing.SetAttribute("name", GenerateUniqueName(string.IsNullOrWhiteSpace(baseName) ? "Shape" : baseName));
            }

            if (shifted) duplicate.SetPosition(left, top);

            parent.InsertBefore(clone, original.NextNode);
            part.MarkEdited();
            ShapeList.Insert(index + 1, duplicate);
            return duplicate;
        }

        private static IEnumerable<OfficeXmlNode> ShapeNodesToRenumber(OfficeXmlNode clone) {
            var pending = new Stack<OfficeXmlNode>();
            pending.Push(clone);
            while (pending.Count > 0) {
                OfficeXmlNode node = pending.Pop();
                // The SDK reserves an ID for each direct child of a group before checking the element type.
                yield return node;
                if (!node.Is(OfficeXmlNamespaces.PresentationML, "grpSp")) continue;
                for (OfficeXmlNode child = node.LastChild; child.Exists; child = child.PreviousSibling) pending.Push(child);
            }
        }

        private static string GroupMemberBaseName(OfficeXmlNode node, OfficeXmlNode drawing) {
            if (drawing.Attribute("name") is { } name) return name;
            if (node.NamespaceId != OfficeXmlNamespaces.PresentationML) return "Shape";
            return node.LocalName switch {
                "cxnSp" => "Connector",
                "pic" => "Picture",
                "graphicFrame" => GraphicFrameBaseName(node),
                "grpSp" => "Group",
                _ => "Shape"
            };
        }

        private static string GraphicFrameBaseName(OfficeXmlNode frame) {
            OfficeXmlNode data = frame.Child(OfficeXmlNamespaces.DrawingML, "graphic").Child(OfficeXmlNamespaces.DrawingML, "graphicData");
            if (data.Child(OfficeXmlNamespaces.DrawingML, "tbl").Exists) return "Table";
            return data.Child(OfficeXmlNamespaces.Chart, "chart").Exists ? "Chart" : "GraphicFrame";
        }

        // The shared clone helper recursively copies all markup, including extension payloads. Unusually deep shapes
        // take the original copying path without recursing through the owned tree.
        private static bool HasBoundedCopyDepth(OfficeXmlNode root) {
            var pending = new Stack<(OfficeXmlNode Node, int Depth)>();
            pending.Push((root, 1));
            while (pending.Count > 0) {
                (OfficeXmlNode node, int depth) = pending.Pop();
                if (depth > MaxPartCloneDepth) return false;
                for (OfficeXmlNode child = node.FirstChild; child.Exists; child = child.NextSibling) pending.Push((child, depth + 1));
            }
            return true;
        }
    }
}
