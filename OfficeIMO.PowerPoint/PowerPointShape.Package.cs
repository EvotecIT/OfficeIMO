using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Shape members read and edited on the slide's package tree while the SDK is not open. Each returns false when it
    /// can't answer exactly as the SDK element would, and the member then binds the shape and reads it through the SDK.
    /// </summary>
    public abstract partial class PowerPointShape {
        /// <summary>The shape's package tree and element while it works without the SDK, or false once it is bound.</summary>
        internal bool TryGetPackageElement(out OfficeXmlTree tree, out int node) {
            tree = PackageTree!;
            node = PackageNode;
            // Once the SDK opened the package, the tree is gone and the shape binds to the SDK element instead.
            return IsPackageBacked && tree != null && OwnerSlide?.OwnerDocument?.IsSdkLoaded == false;
        }

        /// <summary>
        /// The package to edit the shape in: false when the document can't be edited without the SDK, and the member then
        /// edits the shape through it.
        /// </summary>
        internal bool TryEditPackage(out OfficeXmlTree tree, out int node) {
            node = PackageNode;
            tree = PackageTree!;
            return TryGetPackageElement(out tree, out node) && OwnerSlide!.OwnerDocument!.CanEditPackage;
        }

        /// <summary>Records that the shape's package tree changed, so the next save or SDK open writes it.</summary>
        internal void MarkPackageEdited() => OwnerSlide?.OwnerDocument?.MarkPackageEdited();

        // The p:cNvPr of the shape's non-visual properties, as GetNonVisualDrawingProperties finds it without creating it.
        private static int FindNonVisualDrawing(OfficeXmlTree tree, int node) {
            string? container = tree.LocalNameOf(node) switch {
                "sp" => "nvSpPr",
                "cxnSp" => "nvCxnSpPr",
                "pic" => "nvPicPr",
                "graphicFrame" => "nvGraphicFramePr",
                "grpSp" => "nvGrpSpPr",
                _ => null
            };
            if (container == null || tree.NamespaceOf(node) != OfficeXmlNamespaces.PresentationML) return OfficeXmlTree.NoNode;
            int properties = tree.FindChild(node, OfficeXmlNamespaces.PresentationML, container);
            return properties == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : tree.FindChild(properties, OfficeXmlNamespaces.PresentationML, "cNvPr");
        }

        private bool TryReadPackageDrawingAttribute(string name, out string? value) {
            value = null;
            if (!TryGetPackageElement(out OfficeXmlTree tree, out int node)) return false;
            int drawing = FindNonVisualDrawing(tree, node);
            if (drawing != OfficeXmlTree.NoNode) value = tree.GetAttribute(drawing, OfficeXmlNamespaces.None, name);
            return true;
        }

        private bool TryGetPackageId(out uint? id) {
            id = null;
            if (!TryReadPackageDrawingAttribute("id", out string? value)) return false;
            if (value == null) return true;
            if (!uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out uint parsed)) return false;
            id = parsed;
            return true;
        }

        private bool TryGetPackageHidden(out bool hidden) {
            hidden = false;
            if (!TryReadPackageDrawingAttribute("hidden", out string? value)) return false;
            if (value == null) return true;
            switch (value) {
                case "1": case "true": hidden = true; return true;
                case "0": case "false": return true;
                default: return false;
            }
        }

        // The a:off or a:ext of the shape's transform, as PeekOffset and PeekExtents find it.
        private bool TryGetPackageTransformPoint(string point, out int element) {
            element = OfficeXmlTree.NoNode;
            if (!TryGetPackageElement(out OfficeXmlTree tree, out int node) || tree.NamespaceOf(node) != OfficeXmlNamespaces.PresentationML) return false;
            int transform;
            switch (tree.LocalNameOf(node)) {
                case "sp":
                case "cxnSp":
                case "pic": {
                    int properties = tree.FindChild(node, OfficeXmlNamespaces.PresentationML, "spPr");
                    transform = properties == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : tree.FindChild(properties, OfficeXmlNamespaces.DrawingML, "xfrm");
                    break;
                }
                case "graphicFrame":
                    transform = tree.FindChild(node, OfficeXmlNamespaces.PresentationML, "xfrm");
                    break;
                case "grpSp": {
                    int properties = tree.FindChild(node, OfficeXmlNamespaces.PresentationML, "grpSpPr");
                    transform = properties == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : tree.FindChild(properties, OfficeXmlNamespaces.DrawingML, "xfrm");
                    break;
                }
                default:
                    return false;
            }
            if (transform != OfficeXmlTree.NoNode) element = tree.FindChild(transform, OfficeXmlNamespaces.DrawingML, point);
            return true;
        }

        private bool TryGetPackageCoordinate(string point, string attribute, out long value) {
            value = 0L;
            if (!TryGetPackageTransformPoint(point, out int element)) return false;
            if (element == OfficeXmlTree.NoNode) return true;
            string? text = PackageTree!.GetAttribute(element, OfficeXmlNamespaces.None, attribute);
            return text == null || long.TryParse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out value);
        }
    }
}
