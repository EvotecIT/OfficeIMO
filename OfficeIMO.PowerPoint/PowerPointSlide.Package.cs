using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Shapes of a slide whose SDK part is not open, read from the slide's tree in the document's package, as
    /// <see cref="CreateShapeFromElement(DocumentFormat.OpenXml.OpenXmlElement)"/> reads the SDK's shape tree.
    /// </summary>
    public partial class PowerPointSlide {
        /// <summary>The slide's tree in the document's package, or null when the slide has to be read through the SDK.</summary>
        internal OfficeXmlTree? PackageTree => IsPackageBacked && PackagePartName != null ? OwnerDocument?.GetPackageXml(PackagePartName) : null;

        // Loads the shapes from the package tree; false, having loaded nothing, when the slide holds shapes that are
        // read through the SDK.
        private bool TryLoadPackageShapes() {
            if (PackageTree is not { } slide) return false;
            if (!slide.IsElement(slide.Root, OfficeXmlNamespaces.PresentationML, "sld")) return false;
            int data = slide.FindChild(slide.Root, OfficeXmlNamespaces.PresentationML, "cSld");
            int tree = data == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : slide.FindChild(data, OfficeXmlNamespaces.PresentationML, "spTree");
            if (tree == OfficeXmlTree.NoNode) return true;

            var shapes = new List<PowerPointShape>();
            for (int child = slide.FirstChildOf(tree); child != OfficeXmlTree.NoNode; child = slide.NextSiblingOf(child)) {
                if (slide.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
                if (!TryCreatePackageShape(slide, child, out PowerPointShape? shape)) return false;
                if (shape != null) shapes.Add(shape);
            }
            _shapes.AddRange(shapes);
            return true;
        }

        // A shape for an element of the shape tree, null for an element the slide model does not list; false for a shape
        // that is read through the SDK.
        private bool TryCreatePackageShape(OfficeXmlTree slide, int element, out PowerPointShape? shape) {
            shape = null;
            if (slide.IsElement(element, OfficeXmlNamespaces.PresentationML, "sp")) {
                shape = slide.FindChild(element, OfficeXmlNamespaces.PresentationML, "txBody") != OfficeXmlTree.NoNode
                    ? new PowerPointTextBox(this, slide, element)
                    : new PowerPointAutoShape(this, slide, element);
                return true;
            }
            if (slide.IsElement(element, OfficeXmlNamespaces.PresentationML, "pic")) {
                if (IsMedia(slide, element)) return false;
                shape = new PowerPointImage(this, slide, element);
                return true;
            }
            if (slide.IsElement(element, OfficeXmlNamespaces.PresentationML, "graphicFrame")) {
                int graphic = slide.FindChild(element, OfficeXmlNamespaces.DrawingML, "graphic");
                int data = graphic == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : slide.FindChild(graphic, OfficeXmlNamespaces.DrawingML, "graphicData");
                if (data == OfficeXmlTree.NoNode) return true;
                if (slide.FindChild(data, OfficeXmlNamespaces.DrawingML, "tbl") != OfficeXmlTree.NoNode) {
                    shape = new PowerPointTable(this, slide, element);
                    return true;
                }
                // Charts, diagrams and embedded objects are read through the SDK.
                for (int child = slide.FirstChildOf(data); child != OfficeXmlTree.NoNode; child = slide.NextSiblingOf(child)) {
                    if (slide.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
                    string name = slide.LocalNameOf(child);
                    if (name is "chart" or "relIds" or "oleObj") return false;
                }
                return true;
            }
            // Groups and connectors are read through the SDK; anything else is not a shape the model lists.
            return !slide.IsElement(element, OfficeXmlNamespaces.PresentationML, "grpSp")
                && !slide.IsElement(element, OfficeXmlNamespaces.PresentationML, "cxnSp");
        }

        // A picture that plays audio or video, as PowerPointMedia.TryGetMediaKind reads it.
        private static bool IsMedia(OfficeXmlTree slide, int picture) {
            int properties = slide.FindChild(picture, OfficeXmlNamespaces.PresentationML, "nvPicPr");
            int application = properties == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : slide.FindChild(properties, OfficeXmlNamespaces.PresentationML, "nvPr");
            return application != OfficeXmlTree.NoNode
                && (slide.FindChild(application, OfficeXmlNamespaces.DrawingML, "audioFile") != OfficeXmlTree.NoNode
                    || slide.FindChild(application, OfficeXmlNamespaces.DrawingML, "videoFile") != OfficeXmlTree.NoNode);
        }
    }
}
