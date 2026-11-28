using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using D = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointUtils {
        // Drawing and section edits keep comments and processing instructions in order when moving nodes.
        internal static OfficeXmlNode NextNode(this OfficeXmlNode element) =>
            element.Tree is { } tree ? new OfficeXmlNode(tree, tree.NextSiblingOf(element.Node))
                : OfficeXmlNode.FromSdk((element.SdkElement as OpenXmlElement)?.NextSibling());

        internal static List<OfficeXmlNode> ChildNodes(this OfficeXmlNode element) {
            var children = new List<OfficeXmlNode>();
            if (element.Tree is { } tree) {
                for (int child = tree.FirstChildOf(element.Node); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
                    children.Add(new OfficeXmlNode(tree, child));
                }
            } else if (element.SdkElement is OpenXmlElement sdk) {
                foreach (OpenXmlElement child in sdk.ChildElements) children.Add(OfficeXmlNode.FromSdk(child));
            }
            return children;
        }

        internal static void RemoveAllChildren(this OfficeXmlNode element) {
            if (element.Tree is { } tree) {
                for (int child = tree.FirstChildOf(element.Node); child != OfficeXmlTree.NoNode; child = tree.FirstChildOf(element.Node)) {
                    tree.Remove(child);
                }
            } else if (element.SdkElement is OpenXmlElement sdk) {
                sdk.RemoveAllChildren();
            }
        }

        internal static GroupShapeProperties CreateDefaultGroupShapeProperties() {
            return new GroupShapeProperties(CreateDefaultTransformGroup());
        }

        private static D.TransformGroup CreateDefaultTransformGroup() {
            return new D.TransformGroup(
                new D.Offset { X = 0L, Y = 0L },
                new D.Extents { Cx = 0L, Cy = 0L },
                new D.ChildOffset { X = 0L, Y = 0L },
                new D.ChildExtents { Cx = 0L, Cy = 0L });
        }

    }
}
