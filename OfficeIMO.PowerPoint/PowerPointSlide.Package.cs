using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Shapes of a slide whose SDK part is not open, read from the slide's tree in the document's package, as
    /// <see cref="CreateShapeFromElement(DocumentFormat.OpenXml.OpenXmlElement)"/> reads the SDK's shape tree.
    /// </summary>
    public partial class PowerPointSlide {
        /// <summary>The slide's tree in the document's package, or null when the slide has to be read through the SDK.</summary>
        internal OfficeXmlTree? PackageTree => IsPackageBacked && PackagePartName != null ? OwnerDocument?.GetPackageXml(PackagePartName) : null;

        // Whether the shapes were loaded from the package tree, so new shapes can be added there.
        private bool _shapesFromPackage;

        /// <summary>
        /// The slide's shape tree when shapes can be added to it without the SDK: the slide works on the package, its
        /// shapes were read from there, and the document can be edited that way.
        /// </summary>
        private bool TryGetEditablePackageShapeTree(out OfficeXmlTree slide, out int tree) {
            slide = null!;
            tree = OfficeXmlTree.NoNode;
            if (!IsPackageBacked || OwnerDocument?.CanEditPackage != true) return false;
            EnsureShapesLoaded();
            if (!_shapesFromPackage || !IsPackageBacked || PackageTree is not { } package) return false;
            int data = package.FindChild(package.Root, OfficeXmlNamespaces.PresentationML, "cSld");
            tree = data == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : package.FindChild(data, OfficeXmlNamespaces.PresentationML, "spTree");
            slide = package;
            return tree != OfficeXmlTree.NoNode;
        }

        // The largest p:cNvPr id in the shape tree, groups included, as EnsureShapeIdsResolved reads it; false when the
        // slide is read through the SDK or an id is not a number.
        private bool TryGetPackageMaxShapeId(out uint maxId) {
            maxId = 1;
            if (!_shapesFromPackage || !IsPackageBacked || PackageTree is not { } slide) return false;
            int data = slide.FindChild(slide.Root, OfficeXmlNamespaces.PresentationML, "cSld");
            int tree = data == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : slide.FindChild(data, OfficeXmlNamespaces.PresentationML, "spTree");
            if (tree == OfficeXmlTree.NoNode) return true;
            var pending = new Stack<int>();
            pending.Push(tree);
            while (pending.Count > 0) {
                for (int child = slide.FirstChildOf(pending.Pop()); child != OfficeXmlTree.NoNode; child = slide.NextSiblingOf(child)) {
                    if (slide.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
                    if (slide.IsElement(child, OfficeXmlNamespaces.PresentationML, "cNvPr")) {
                        string? value = slide.GetAttribute(child, OfficeXmlNamespaces.None, "id");
                        uint id = 0U;
                        if (value != null && !uint.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out id)) return false;
                        if (id > maxId) maxId = id;
                    }
                    pending.Push(child);
                }
            }
            return true;
        }

        // Adds a text box to the package tree as AddTextBox writes it; null when the slide or document can't be edited
        // that way, and the caller adds it through the SDK.
        private PowerPointTextBox? TryAddPackageTextBox(string text, long left, long top, long width, long height) {
            if (!TryGetEditablePackageShapeTree(out OfficeXmlTree slide, out int tree) || !TryGetPackageMaxShapeId(out _)) return null;
            string name = GenerateUniqueName("TextBox");
            uint id = AllocateShapeId();
            // Reading names or ids can open the SDK for a shape the package does not read; the SDK then adds it.
            if (!IsPackageBacked) return null;

            int shape = slide.CreateElement(OfficeXmlNamespaces.PresentationML, "sp");
            int nonVisual = PowerPointDocument.Append(slide, shape, OfficeXmlNamespaces.PresentationML, "nvSpPr");
            int drawing = PowerPointDocument.Append(slide, nonVisual, OfficeXmlNamespaces.PresentationML, "cNvPr");
            slide.SetAttribute(drawing, OfficeXmlNamespaces.None, "id", id.ToString(System.Globalization.CultureInfo.InvariantCulture));
            slide.SetAttribute(drawing, OfficeXmlNamespaces.None, "name", name);
            int shapeDrawing = PowerPointDocument.Append(slide, nonVisual, OfficeXmlNamespaces.PresentationML, "cNvSpPr");
            slide.SetAttribute(shapeDrawing, OfficeXmlNamespaces.None, "txBox", "1");
            slide.SetAttribute(PowerPointDocument.Append(slide, shapeDrawing, OfficeXmlNamespaces.DrawingML, "spLocks"), OfficeXmlNamespaces.None, "noGrp", "1");
            PowerPointDocument.Append(slide, nonVisual, OfficeXmlNamespaces.PresentationML, "nvPr");
            AppendRectangle(slide, PowerPointDocument.Append(slide, shape, OfficeXmlNamespaces.PresentationML, "spPr"), left, top, width, height);
            int body = PowerPointDocument.Append(slide, shape, OfficeXmlNamespaces.PresentationML, "txBody");
            PowerPointDocument.Append(slide, body, OfficeXmlNamespaces.DrawingML, "bodyPr");
            PowerPointDocument.Append(slide, body, OfficeXmlNamespaces.DrawingML, "lstStyle");
            PowerPointTreeText.AppendInlineContent(slide, PowerPointDocument.Append(slide, body, OfficeXmlNamespaces.DrawingML, "p"), text, OfficeXmlTree.NoNode);
            slide.Append(tree, shape);
            OwnerDocument!.MarkPackageEdited();
            return TrackShape(new PowerPointTextBox(this, slide, shape));
        }

        // Adds a picture to the package tree as AddImage writes it, its image part shared with an identical image added
        // before; null when the slide, the document or the image name is left to the SDK.
        private PowerPointImage? TryAddPackageImage(byte[] data, OfficeIMO.Drawing.OfficeImageFormat imageType, string extension,
            long left, long top, long width, long height) {
            if (imageType == OfficeIMO.Drawing.OfficeImageFormat.Svg || OwnerDocument is not { } document
                || !TryGetEditablePackageShapeTree(out OfficeXmlTree slide, out int tree) || !TryGetPackageMaxShapeId(out _)
                || !document.TryAddPackageImagePart(PackagePartName!, data, imageType, extension, out string relationshipId, out int pixelWidth, out int pixelHeight)) {
                return null;
            }
            uint id = AllocateShapeId();
            string name = GenerateUniqueName("Picture");
            if (!IsPackageBacked) throw new InvalidOperationException("The slide opened the SDK while a picture was added to its package.");

            // The picture keeps the image's aspect ratio inside the box, as ContainInFrame fits it.
            if (pixelWidth > 0 && pixelHeight > 0 && width > 0 && height > 0) {
                double imageAspect = pixelWidth / (double)pixelHeight, boxAspect = width / (double)height;
                if (Math.Abs(imageAspect / ((width / 12700D) / (height / 12700D)) - 1D) >= 0.005D) {
                    double newWidth, newHeight;
                    if (imageAspect > boxAspect) {
                        newWidth = width;
                        newHeight = width / imageAspect;
                    } else {
                        newHeight = height;
                        newWidth = height * imageAspect;
                    }
                    left += (long)Math.Round((width - newWidth) / 2);
                    top += (long)Math.Round((height - newHeight) / 2);
                    width = (long)Math.Round(newWidth);
                    height = (long)Math.Round(newHeight);
                }
            }

            int picture = slide.CreateElement(OfficeXmlNamespaces.PresentationML, "pic");
            int nonVisual = PowerPointDocument.Append(slide, picture, OfficeXmlNamespaces.PresentationML, "nvPicPr");
            int drawing = PowerPointDocument.Append(slide, nonVisual, OfficeXmlNamespaces.PresentationML, "cNvPr");
            slide.SetAttribute(drawing, OfficeXmlNamespaces.None, "id", id.ToString(System.Globalization.CultureInfo.InvariantCulture));
            slide.SetAttribute(drawing, OfficeXmlNamespaces.None, "name", name);
            int pictureDrawing = PowerPointDocument.Append(slide, nonVisual, OfficeXmlNamespaces.PresentationML, "cNvPicPr");
            slide.SetAttribute(PowerPointDocument.Append(slide, pictureDrawing, OfficeXmlNamespaces.DrawingML, "picLocks"), OfficeXmlNamespaces.None, "noChangeAspect", "1");
            PowerPointDocument.Append(slide, nonVisual, OfficeXmlNamespaces.PresentationML, "nvPr");
            int fill = PowerPointDocument.Append(slide, picture, OfficeXmlNamespaces.PresentationML, "blipFill");
            slide.SetAttribute(PowerPointDocument.Append(slide, fill, OfficeXmlNamespaces.DrawingML, "blip"), OfficeXmlNamespaces.Relationships, "embed", relationshipId);
            PowerPointDocument.Append(slide, PowerPointDocument.Append(slide, fill, OfficeXmlNamespaces.DrawingML, "stretch"), OfficeXmlNamespaces.DrawingML, "fillRect");
            AppendRectangle(slide, PowerPointDocument.Append(slide, picture, OfficeXmlNamespaces.PresentationML, "spPr"), left, top, width, height);
            slide.Append(tree, picture);
            document.MarkPackageEdited();
            return TrackShape(new PowerPointImage(this, slide, picture));
        }

        // A transform and a rectangle preset geometry, as the model writes them into new shapes.
        private static void AppendRectangle(OfficeXmlTree slide, int properties, long left, long top, long width, long height) {
            int transform = PowerPointDocument.Append(slide, properties, OfficeXmlNamespaces.DrawingML, "xfrm");
            int offset = PowerPointDocument.Append(slide, transform, OfficeXmlNamespaces.DrawingML, "off");
            slide.SetAttribute(offset, OfficeXmlNamespaces.None, "x", left.ToString(System.Globalization.CultureInfo.InvariantCulture));
            slide.SetAttribute(offset, OfficeXmlNamespaces.None, "y", top.ToString(System.Globalization.CultureInfo.InvariantCulture));
            int extents = PowerPointDocument.Append(slide, transform, OfficeXmlNamespaces.DrawingML, "ext");
            slide.SetAttribute(extents, OfficeXmlNamespaces.None, "cx", width.ToString(System.Globalization.CultureInfo.InvariantCulture));
            slide.SetAttribute(extents, OfficeXmlNamespaces.None, "cy", height.ToString(System.Globalization.CultureInfo.InvariantCulture));
            int geometry = PowerPointDocument.Append(slide, properties, OfficeXmlNamespaces.DrawingML, "prstGeom");
            slide.SetAttribute(geometry, OfficeXmlNamespaces.None, "prst", "rect");
            PowerPointDocument.Append(slide, geometry, OfficeXmlNamespaces.DrawingML, "avLst");
        }

        // Loads the shapes from the package tree; false, having loaded nothing, when the slide holds shapes that are
        // read through the SDK.
        private bool TryLoadPackageShapes() {
            if (PackageTree is not { } slide) return false;
            if (!slide.IsElement(slide.Root, OfficeXmlNamespaces.PresentationML, "sld")) return false;
            int data = slide.FindChild(slide.Root, OfficeXmlNamespaces.PresentationML, "cSld");
            int tree = data == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : slide.FindChild(data, OfficeXmlNamespaces.PresentationML, "spTree");
            if (tree == OfficeXmlTree.NoNode) {
                _shapesFromPackage = true;
                return true;
            }

            var shapes = new List<PowerPointShape>();
            for (int child = slide.FirstChildOf(tree); child != OfficeXmlTree.NoNode; child = slide.NextSiblingOf(child)) {
                if (slide.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
                if (!TryCreatePackageShape(slide, child, out PowerPointShape? shape)) return false;
                if (shape != null) shapes.Add(shape);
            }
            _shapes.AddRange(shapes);
            _shapesFromPackage = true;
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
