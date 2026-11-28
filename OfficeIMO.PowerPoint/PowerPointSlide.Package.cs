using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Shapes of a slide whose SDK part is not open, read from the slide's tree in the document's package, as
    /// CreateShapeFromElement reads the SDK's shape tree.
    /// </summary>
    public partial class PowerPointSlide {
        /// <summary>The slide's tree in the document's package, or null when the slide has to be read through the SDK.</summary>
        internal OfficeXmlTree? PackageTree => IsPackageBacked && !RemovedFromPackage && PackagePartName != null ? OwnerDocument?.GetPackageXml(PackagePartName) : null;

        /// <summary>The slide's part: in the document's package while it works there, its SDK part otherwise.</summary>
        internal PowerPointPartRef PartRef => IsPackageBacked && !RemovedFromPackage && PackagePartName != null && OwnerDocument is { IsSdkLoaded: false } document
            ? PowerPointPartRef.Package(document, PackagePartName)
            : PowerPointPartRef.FromSdk(SdkSlidePart);

        // Whether the shapes were loaded from the package tree, so new shapes can be added there.
        private bool _shapesFromPackage;

        internal bool HasMaterializedPackageShapes => _shapesFromPackage && _shapes.Count != 0;

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
                    string name = slide.LocalNameOf(child);
                    // A chosen compatibility branch can expose its drawing IDs after the package is saved.
                    if (name == "extLst") continue;
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

        // Adds a picture to the package tree as AddImage writes it, its image part shared with an identical image added
        // before; null when the slide, the document or the image name is left to the SDK.
        private PowerPointImage? TryAddPackageImage(byte[] data, OfficeIMO.Drawing.OfficeImageFormat imageType, string contentType, string extension,
            long left, long top, long width, long height) {
            if (imageType == OfficeIMO.Drawing.OfficeImageFormat.Svg || OwnerDocument is not { } document
                || !TryGetEditablePackageShapeTree(out OfficeXmlTree slide, out int tree) || !TryGetPackageMaxShapeId(out uint maxShapeId)
                // A slide out of shape ids makes the model fail before it adds the image part, so nothing is added here.
                || maxShapeId == uint.MaxValue
                || !document.TryAddPackageImagePart(PackagePartName!, data, contentType, extension, out string relationshipId, out int pixelWidth, out int pixelHeight)) {
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

        // Adds a table to the package tree as AddTable writes it with explicit bounds: the default style, columns and rows
        // shared out evenly, an empty en-US run in each cell. Null when the slide or document is left to the SDK.
        private PowerPointTable? TryAddPackageTable(int rows, int columns, long left, long top, long width, long height) {
            if (!TryGetEditablePackageShapeTree(out OfficeXmlTree slide, out int tree) || !TryGetPackageMaxShapeId(out _)) return null;
            const string a16 = "http://schemas.microsoft.com/office/drawing/2014/main";
            int a16Namespace = slide.Namespaces.GetId(a16);
            string name = GenerateUniqueName("Table");
            uint id = AllocateShapeId();

            int frame = slide.CreateElement(OfficeXmlNamespaces.PresentationML, "graphicFrame");
            int nonVisual = PowerPointDocument.Append(slide, frame, OfficeXmlNamespaces.PresentationML, "nvGraphicFramePr");
            int drawing = PowerPointDocument.Append(slide, nonVisual, OfficeXmlNamespaces.PresentationML, "cNvPr");
            slide.SetAttribute(drawing, OfficeXmlNamespaces.None, "id", id.ToString(System.Globalization.CultureInfo.InvariantCulture));
            slide.SetAttribute(drawing, OfficeXmlNamespaces.None, "name", name);
            PowerPointDocument.Append(slide, nonVisual, OfficeXmlNamespaces.PresentationML, "cNvGraphicFramePr");
            PowerPointDocument.Append(slide, nonVisual, OfficeXmlNamespaces.PresentationML, "nvPr");
            int transform = PowerPointDocument.Append(slide, frame, OfficeXmlNamespaces.PresentationML, "xfrm");
            SetPoint(slide, PowerPointDocument.Append(slide, transform, OfficeXmlNamespaces.DrawingML, "off"), "x", left, "y", top);
            SetPoint(slide, PowerPointDocument.Append(slide, transform, OfficeXmlNamespaces.DrawingML, "ext"), "cx", width, "cy", height);
            int data = PowerPointDocument.Append(slide, PowerPointDocument.Append(slide, frame, OfficeXmlNamespaces.DrawingML, "graphic"), OfficeXmlNamespaces.DrawingML, "graphicData");
            slide.SetAttribute(data, OfficeXmlNamespaces.None, "uri", "http://schemas.openxmlformats.org/drawingml/2006/table");
            int table = PowerPointDocument.Append(slide, data, OfficeXmlNamespaces.DrawingML, "tbl");
            int properties = PowerPointDocument.Append(slide, table, OfficeXmlNamespaces.DrawingML, "tblPr");
            slide.SetAttribute(properties, OfficeXmlNamespaces.None, "firstRow", "1");
            slide.SetAttribute(properties, OfficeXmlNamespaces.None, "bandRow", "1");
            slide.Append(PowerPointDocument.Append(slide, properties, OfficeXmlNamespaces.DrawingML, "tableStyleId"), slide.CreateText("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"));

            // Widths and heights as SetColumnWidthsEvenly and SetRowHeightsEvenly leave them: the remainder on the last.
            int grid = PowerPointDocument.Append(slide, table, OfficeXmlNamespaces.DrawingML, "tblGrid");
            long assignedWidth = 0;
            for (int column = 0; column < columns; column++) {
                long columnWidth = (long)Math.Round(width * (1D / columns));
                assignedWidth += columnWidth;
                if (column == columns - 1) columnWidth += width - assignedWidth;
                int gridColumn = PowerPointDocument.Append(slide, grid, OfficeXmlNamespaces.DrawingML, "gridCol");
                slide.SetAttribute(gridColumn, OfficeXmlNamespaces.None, "w", columnWidth.ToString(System.Globalization.CultureInfo.InvariantCulture));
                AppendTableExtension(slide, gridColumn, "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}", a16Namespace, a16, "colId", 20000U + (uint)column);
            }
            long rowHeight = (long)Math.Floor(height / (double)rows);
            for (int row = 0; row < rows; row++) {
                int tableRow = PowerPointDocument.Append(slide, table, OfficeXmlNamespaces.DrawingML, "tr");
                long thisHeight = row == rows - 1 ? height - rowHeight * (rows - 1) : rowHeight;
                slide.SetAttribute(tableRow, OfficeXmlNamespaces.None, "h", thisHeight.ToString(System.Globalization.CultureInfo.InvariantCulture));
                for (int column = 0; column < columns; column++) {
                    int cell = PowerPointDocument.Append(slide, tableRow, OfficeXmlNamespaces.DrawingML, "tc");
                    int body = PowerPointDocument.Append(slide, cell, OfficeXmlNamespaces.DrawingML, "txBody");
                    PowerPointDocument.Append(slide, body, OfficeXmlNamespaces.DrawingML, "bodyPr");
                    PowerPointDocument.Append(slide, body, OfficeXmlNamespaces.DrawingML, "lstStyle");
                    int paragraph = PowerPointDocument.Append(slide, body, OfficeXmlNamespaces.DrawingML, "p");
                    int run = PowerPointDocument.Append(slide, paragraph, OfficeXmlNamespaces.DrawingML, "r");
                    slide.SetAttribute(PowerPointDocument.Append(slide, run, OfficeXmlNamespaces.DrawingML, "rPr"), OfficeXmlNamespaces.None, "lang", PowerPointTableTextDefaults.Language);
                    PowerPointDocument.Append(slide, run, OfficeXmlNamespaces.DrawingML, "t");
                    slide.SetAttribute(PowerPointDocument.Append(slide, paragraph, OfficeXmlNamespaces.DrawingML, "endParaRPr"), OfficeXmlNamespaces.None, "lang", PowerPointTableTextDefaults.Language);
                    PowerPointDocument.Append(slide, cell, OfficeXmlNamespaces.DrawingML, "tcPr");
                }
                AppendTableExtension(slide, tableRow, "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}", a16Namespace, a16, "rowId", 10000U + (uint)row);
            }
            slide.Append(tree, frame);
            OwnerDocument!.MarkPackageEdited();
            return TrackShape(new PowerPointTable(this, slide, frame));
        }

        private static void AppendTableExtension(OfficeXmlTree slide, int parent, string uri, int a16Namespace, string a16, string localName, uint value) {
            int extension = PowerPointDocument.Append(slide, PowerPointDocument.Append(slide, parent, OfficeXmlNamespaces.DrawingML, "extLst"), OfficeXmlNamespaces.DrawingML, "ext");
            slide.SetAttribute(extension, OfficeXmlNamespaces.None, "uri", uri);
            int element = slide.CreateElement(a16Namespace, localName);
            slide.DeclareNamespace(element, "a16", a16);
            slide.SetAttribute(element, OfficeXmlNamespaces.None, "val", value.ToString(System.Globalization.CultureInfo.InvariantCulture));
            slide.Append(extension, element);
        }

        private static void SetPoint(OfficeXmlTree slide, int element, string first, long firstValue, string second, long secondValue) {
            slide.SetAttribute(element, OfficeXmlNamespaces.None, first, firstValue.ToString(System.Globalization.CultureInfo.InvariantCulture));
            slide.SetAttribute(element, OfficeXmlNamespaces.None, second, secondValue.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        // Adds an auto shape to the package tree as AddShape writes it; null when the slide or document is left to the SDK.
        private PowerPointAutoShape? TryAddPackageShape(OfficePresetShapeType shapeType, string preset, long left, long top, long width, long height, string? name) {
            if (!TryGetEditablePackageShapeTree(out OfficeXmlTree slide, out int tree) || !TryGetPackageMaxShapeId(out _)) return null;
            string shapeName = name ?? GenerateUniqueName(shapeType.ToString());
            uint id = AllocateShapeId();
            int shape = slide.CreateElement(OfficeXmlNamespaces.PresentationML, "sp");
            int nonVisual = PowerPointDocument.Append(slide, shape, OfficeXmlNamespaces.PresentationML, "nvSpPr");
            int drawing = PowerPointDocument.Append(slide, nonVisual, OfficeXmlNamespaces.PresentationML, "cNvPr");
            slide.SetAttribute(drawing, OfficeXmlNamespaces.None, "id", id.ToString(System.Globalization.CultureInfo.InvariantCulture));
            slide.SetAttribute(drawing, OfficeXmlNamespaces.None, "name", shapeName);
            int shapeDrawing = PowerPointDocument.Append(slide, nonVisual, OfficeXmlNamespaces.PresentationML, "cNvSpPr");
            slide.SetAttribute(PowerPointDocument.Append(slide, shapeDrawing, OfficeXmlNamespaces.DrawingML, "spLocks"), OfficeXmlNamespaces.None, "noGrp", "1");
            PowerPointDocument.Append(slide, nonVisual, OfficeXmlNamespaces.PresentationML, "nvPr");
            int properties = PowerPointDocument.Append(slide, shape, OfficeXmlNamespaces.PresentationML, "spPr");
            int transform = PowerPointDocument.Append(slide, properties, OfficeXmlNamespaces.DrawingML, "xfrm");
            SetPoint(slide, PowerPointDocument.Append(slide, transform, OfficeXmlNamespaces.DrawingML, "off"), "x", left, "y", top);
            SetPoint(slide, PowerPointDocument.Append(slide, transform, OfficeXmlNamespaces.DrawingML, "ext"), "cx", width, "cy", height);
            int geometry = PowerPointDocument.Append(slide, properties, OfficeXmlNamespaces.DrawingML, "prstGeom");
            slide.SetAttribute(geometry, OfficeXmlNamespaces.None, "prst", preset);
            PowerPointDocument.Append(slide, geometry, OfficeXmlNamespaces.DrawingML, "avLst");
            if (shapeType == OfficePresetShapeType.Line) PowerPointDocument.Append(slide, properties, OfficeXmlNamespaces.DrawingML, "noFill");
            slide.Append(tree, shape);
            OwnerDocument!.MarkPackageEdited();
            return TrackShape(new PowerPointAutoShape(this, slide, shape));
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

        // Wraps a drawing child while retaining the part that owns its relationships and the slide that renders it.
        private PowerPointShape? WrapSlideShape(OfficeXmlNode element, PowerPointPartRef owner) {
            if (element.Tree is { } tree) {
                if (owner.Equals(PartRef)) {
                    TryCreatePackageShape(tree, element.Node, out PowerPointShape? slideShape);
                    return slideShape;
                }

                return PowerPointMasterSurface.CreatePackageShape(owner, element)?.AttachToPackageSlide(this);
            }

            return WrapSdkSlideShape(element, owner);
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
                shape = MediaKind(slide, element) is { } kind ? new PowerPointMedia(this, slide, element, kind) : new PowerPointImage(this, slide, element);
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
                // Charts, diagrams and embedded objects, in the order the model checks them; each reads its content
                // through the SDK when used.
                if (slide.FindChild(data, OfficeXmlNamespaces.Chart, "chart") != OfficeXmlTree.NoNode) {
                    shape = new PowerPointChart(this, slide, element);
                } else if (slide.FindChild(data, OfficeXmlNamespaces.Diagram, "relIds") != OfficeXmlTree.NoNode) {
                    shape = new PowerPointSmartArt(this, slide, element);
                } else if (slide.FindChild(data, OfficeXmlNamespaces.PresentationML, "oleObj") != OfficeXmlTree.NoNode) {
                    shape = new PowerPointOleObject(this, slide, element);
                }
                return true;
            }
            if (slide.IsElement(element, OfficeXmlNamespaces.PresentationML, "grpSp")) {
                shape = new PowerPointGroupShape(this, slide, element);
                return true;
            }
            if (slide.IsElement(element, OfficeXmlNamespaces.PresentationML, "cxnSp")) {
                shape = new PowerPointConnectionShape(this, slide, element);
                return true;
            }
            // Anything else is not a shape the model lists.
            return true;
        }

        // The kind of media a picture plays, as PowerPointMedia.TryGetMediaKind reads it; null for a still picture.
        private static PowerPointMediaKind? MediaKind(OfficeXmlTree slide, int picture) {
            int properties = slide.FindChild(picture, OfficeXmlNamespaces.PresentationML, "nvPicPr");
            int application = properties == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : slide.FindChild(properties, OfficeXmlNamespaces.PresentationML, "nvPr");
            if (application == OfficeXmlTree.NoNode) return null;
            if (slide.FindChild(application, OfficeXmlNamespaces.DrawingML, "audioFile") != OfficeXmlTree.NoNode) return PowerPointMediaKind.Audio;
            if (slide.FindChild(application, OfficeXmlNamespaces.DrawingML, "videoFile") != OfficeXmlTree.NoNode) return PowerPointMediaKind.Video;
            return null;
        }
    }
}
