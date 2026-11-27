using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Shared editing surface of a slide master and a slide layout. Shapes, placeholders and the background
    ///     placed here are inherited by every slide that uses the master or layout.
    /// </summary>
    public abstract class PowerPointMasterSurface {
        private const int P = OfficeXmlNamespaces.PresentationML;
        private const int D = OfficeXmlNamespaces.DrawingML;
        private bool _removed;

        internal PowerPointMasterSurface(PowerPointDocument document, PowerPointPartRef part) {
            Document = document;
            _part = part;
        }

        private readonly PowerPointPartRef _part;

        internal PowerPointDocument Document { get; }

        /// <summary>The part's name, which identifies the master or layout.</summary>
        internal string PartName => _part.Name;

        /// <summary>The master's or layout's part as it is held now.</summary>
        internal PowerPointPartRef PartRef => _part.Current;

        /// <summary>The part's root element (<c>p:sldMaster</c> or <c>p:sldLayout</c>).</summary>
        internal OfficeXmlNode Root => PartRef.Root;

        /// <summary>The children of the root element, in sequence order.</summary>
        private protected abstract OfficeXmlChildOrder RootOrder { get; }

        private void MarkEdited() => PartRef.MarkEdited();

        /// <summary>The common slide data, added with a shape tree when missing.</summary>
        internal OfficeXmlNode CommonSlideData {
            get {
                OfficeXmlNode root = Root;
                OfficeXmlNode data = root.Child(P, "cSld");
                if (!data.IsNone) return data;
                data = root.AddOrdered(P, "cSld", RootOrder);
                data.Append(data.CreateElement(P, "spTree"));
                MarkEdited();
                return data;
            }
        }

        /// <summary>Display name shown in PowerPoint's Slide Master view.</summary>
        public string Name {
            get {
                ThrowIfRemoved();
                return CommonSlideData.Attribute("name") ?? string.Empty;
            }
            set {
                ThrowIfRemoved();
                ArgumentNullException.ThrowIfNull(value);
                CommonSlideData.SetAttribute("name", value);
                MarkEdited();
            }
        }

        /// <summary>Shapes drawn on this master or layout, in drawing order.</summary>
        public IReadOnlyList<PowerPointShape> Shapes {
            get {
                ThrowIfRemoved();
                OfficeXmlNode tree = ShapeTree;
                var shapes = new List<PowerPointShape>();
                if (tree.Tree == null || ContainsChart(tree)) {
                    // Charts on a master or layout read through the SDK, as every shape of an SDK-held part does.
                    return Document.CreateSdkSurfaceShapes(this);
                }
                foreach (OfficeXmlNode element in tree.Children()) {
                    if (CreatePackageShape(element) is { } shape) shapes.Add(shape);
                }

                return shapes;
            }
        }

        // Whether the shape tree holds a chart frame, which only reads through the SDK.
        private static bool ContainsChart(OfficeXmlNode tree) {
            foreach (OfficeXmlNode frame in tree.Children(P, "graphicFrame")) {
                if (!frame.Child(D, "graphic").Child(D, "graphicData").Child(OfficeXmlNamespaces.Chart, "chart").IsNone) return true;
            }
            return false;
        }

        // A shape wrapper for an element of the part's package tree, as the SDK model lists a master's or layout's shapes:
        // diagrams and embedded objects only belong to slides.
        private PowerPointShape? CreatePackageShape(OfficeXmlNode element) {
            if (element.NamespaceId != P) return null;
            PowerPointPartRef part = _part;
            OfficeXmlTree tree = element.Tree!;
            int node = element.Node;
            switch (element.LocalName) {
                case "sp":
                    return element.Child(P, "txBody").IsNone ? new PowerPointAutoShape(part, tree, node) : new PowerPointTextBox(part, tree, node);
                case "cxnSp":
                    return new PowerPointConnectionShape(part, tree, node);
                case "pic":
                    return new PowerPointImage(part, tree, node);
                case "grpSp":
                    return new PowerPointGroupShape(part, tree, node);
                case "graphicFrame":
                    return element.Child(D, "graphic").Child(D, "graphicData").Child(D, "tbl").IsNone ? null : new PowerPointTable(part, tree, node);
                default:
                    return null;
            }
        }

        /// <summary>The wrapper for a shape element of this surface.</summary>
        internal PowerPointShape WrapShape(OfficeXmlNode element) =>
            element.Tree != null ? CreatePackageShape(element)! : Document.CreateSdkSurfaceShape(this, element);

        /// <summary>Placeholders defined on this master or layout.</summary>
        public IReadOnlyList<PowerPointTextBox> Placeholders =>
            Shapes.OfType<PowerPointTextBox>().Where(shape => shape.IsPlaceholder).ToList();

        /// <summary>Returns the first placeholder of the given type and, optionally, index.</summary>
        public PowerPointTextBox? GetPlaceholder(PowerPointPlaceholderType placeholderType, uint? index = null) =>
            Placeholders.FirstOrDefault(shape => shape.PlaceholderType == placeholderType
                && (index == null || shape.PlaceholderIndex == index));

        /// <summary>
        ///     Adds a placeholder that slides using this layout (or master) inherit: position, size and text
        ///     formatting. Title placeholders have no index; other types get the next free index when none is given.
        ///     Slide number and date placeholders carry their native fields.
        /// </summary>
        public PowerPointTextBox AddPlaceholder(PowerPointPlaceholderType placeholderType, PowerPointLayoutBox bounds,
            uint? index = null, string? name = null) {
            ThrowIfRemoved();
            if (GetPlaceholder(placeholderType, index) != null && (index != null || IsSingleton(placeholderType))) {
                throw new InvalidOperationException("A " + placeholderType + " placeholder"
                    + (index == null ? string.Empty : " with index " + index.Value) + " already exists on '" + Name + "'.");
            }

            string typeToken = placeholderType.ToOpenXmlToken();
            uint? resolvedIndex = IsTitle(placeholderType) ? null : index ?? NextPlaceholderIndex();
            OfficeXmlNode tree = ShapeTree;
            OfficeXmlNode shape = CreateShape(tree, name ?? UniqueName(placeholderType + " Placeholder"), out OfficeXmlNode application);
            shape.Child(P, "nvSpPr").Child(P, "cNvSpPr").Append(shape.CreateElement(D, "spLocks")).SetAttribute("noGrp", "1");
            OfficeXmlNode placeholder = application.Append(application.CreateElement(P, "ph"));
            placeholder.SetAttribute("type", typeToken);
            placeholder.SetUInt32Attribute("idx", resolvedIndex);
            OfficeXmlNode properties = shape.Append(shape.CreateElement(P, "spPr"));
            properties.Append(CreateTransform(properties, bounds.Left, bounds.Top, bounds.Width, bounds.Height));
            properties.Append(CreatePresetGeometry(properties, "rect"));
            OfficeXmlNode body = shape.Append(shape.CreateElement(P, "txBody"));
            body.Append(body.CreateElement(D, "bodyPr"));
            body.Append(body.CreateElement(D, "lstStyle"));
            body.Append(CreatePlaceholderParagraph(body, placeholderType));
            shape = tree.Append(shape);
            MarkEdited();
            return (PowerPointTextBox)WrapShape(shape);
        }

        /// <summary>Adds an auto shape, such as a band or rule, drawn behind slide content on every slide.</summary>
        public PowerPointAutoShape AddShape(OfficePresetShapeType shapeType, OfficeLength left, OfficeLength top,
            OfficeLength width, OfficeLength height, string? name = null) {
            ThrowIfRemoved();
            string preset = shapeType.ToOpenXmlToken();
            OfficeXmlNode tree = ShapeTree;
            OfficeXmlNode shape = CreateShape(tree, name ?? UniqueName(shapeType.ToString()), out _);
            OfficeXmlNode properties = shape.Append(shape.CreateElement(P, "spPr"));
            properties.Append(CreateTransform(properties, left, top, width, height));
            properties.Append(CreatePresetGeometry(properties, preset));
            shape = tree.Append(shape);
            MarkEdited();
            return (PowerPointAutoShape)WrapShape(shape);
        }

        /// <summary>Adds fixed text, such as a confidentiality notice, shown on every slide.</summary>
        public PowerPointTextBox AddTextBox(string text, OfficeLength left, OfficeLength top, OfficeLength width,
            OfficeLength height) {
            ThrowIfRemoved();
            ArgumentNullException.ThrowIfNull(text);
            OfficeXmlNode tree = ShapeTree;
            OfficeXmlNode shape = CreateShape(tree, UniqueName("TextBox"), out _);
            shape.Child(P, "nvSpPr").Child(P, "cNvSpPr").SetAttribute("txBox", "1");
            OfficeXmlNode properties = shape.Append(shape.CreateElement(P, "spPr"));
            properties.Append(CreateTransform(properties, left, top, width, height));
            properties.Append(CreatePresetGeometry(properties, "rect"));
            properties.Append(properties.CreateElement(D, "noFill"));
            OfficeXmlNode body = shape.Append(shape.CreateElement(P, "txBody"));
            body.Append(body.CreateElement(D, "bodyPr")).SetAttribute("wrap", "square");
            body.Append(body.CreateElement(D, "lstStyle"));
            OfficeXmlNode paragraph = body.Append(body.CreateElement(D, "p"));
            PowerPointTextContent.InsertInlineContent(paragraph, text, default, default);
            shape = tree.Append(shape);
            MarkEdited();
            return (PowerPointTextBox)WrapShape(shape);
        }

        /// <summary>Adds a picture, such as a logo, shown on every slide. The picture keeps its aspect ratio inside the box.</summary>
        public PowerPointImage AddImage(string imagePath, OfficeLength left, OfficeLength top, OfficeLength width,
            OfficeLength height) {
            ArgumentNullException.ThrowIfNull(imagePath);
            if (!File.Exists(imagePath)) throw new FileNotFoundException("Image file not found.", imagePath);
            using FileStream stream = new(imagePath, FileMode.Open, FileAccess.Read);
            return AddImage(stream, PowerPointImageFormatExtensions.FromImagePath(imagePath), left, top, width, height);
        }

        /// <summary>Adds a picture from a stream, shown on every slide. The picture keeps its aspect ratio inside the box.</summary>
        public PowerPointImage AddImage(Stream image, OfficeImageFormat imageType, OfficeLength left, OfficeLength top,
            OfficeLength width, OfficeLength height) {
            ThrowIfRemoved();
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width.ToEmus(), nameof(width));
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height.ToEmus(), nameof(height));
            string relationshipId = EmbedImage(image, imageType);
            OfficeXmlNode tree = ShapeTree;
            OfficeXmlNode picture = tree.CreateElement(P, "pic");
            OfficeXmlNode nonVisual = picture.Append(picture.CreateElement(P, "nvPicPr"));
            AppendDrawingProperties(nonVisual, NextShapeId(), UniqueName("Picture"));
            nonVisual.Append(nonVisual.CreateElement(P, "cNvPicPr")).Append(nonVisual.CreateElement(D, "picLocks")).SetAttribute("noChangeAspect", "1");
            nonVisual.Append(nonVisual.CreateElement(P, "nvPr"));
            OfficeXmlNode fill = picture.Append(picture.CreateElement(P, "blipFill"));
            fill.Append(fill.CreateElement(D, "blip")).SetAttribute(OfficeXmlNamespaces.Relationships, "embed", relationshipId);
            fill.Append(fill.CreateElement(D, "stretch")).Append(fill.CreateElement(D, "fillRect"));
            OfficeXmlNode properties = picture.Append(picture.CreateElement(P, "spPr"));
            properties.Append(CreateTransform(properties, left, top, width, height));
            properties.Append(CreatePresetGeometry(properties, "rect"));
            picture = tree.Append(picture);
            MarkEdited();
            var result = (PowerPointImage)WrapShape(picture);
            result.ContainInFrame();
            return result;
        }

        /// <summary>
        ///     Gets or sets a solid background. Setting <see langword="null" /> removes the local background, so a
        ///     layout falls back to its master's background.
        /// </summary>
        public OfficeColor? BackgroundColor {
            get {
                ThrowIfRemoved();
                return PowerPointDrawingColor.Read(CommonSlideData.Child(P, "bg").Child(P, "bgPr").Child(D, "solidFill"));
            }
            set {
                ThrowIfRemoved();
                if (value == null) {
                    ClearBackground();
                    return;
                }

                OfficeTheme? theme = PowerPointDrawingColor.NeedsTheme(value) ? PowerPointDrawingColor.ThemeFor(PartRef, null) : null;
                OfficeXmlNode data = CommonSlideData;
                SetBackgroundFill(PowerPointDrawingColor.CreateFill(data, value.Value, theme), null);
            }
        }

        /// <summary>Sets a stretched background picture.</summary>
        public void SetBackgroundImage(string imagePath) {
            ArgumentNullException.ThrowIfNull(imagePath);
            if (!File.Exists(imagePath)) throw new FileNotFoundException("Image file not found.", imagePath);
            using FileStream stream = new(imagePath, FileMode.Open, FileAccess.Read);
            SetBackgroundImage(stream, PowerPointImageFormatExtensions.FromImagePath(imagePath));
        }

        /// <summary>Sets a stretched background picture from a stream.</summary>
        public void SetBackgroundImage(Stream image, OfficeImageFormat imageType) {
            ThrowIfRemoved();
            string relationshipId = EmbedImage(image, imageType);
            OfficeXmlNode data = CommonSlideData;
            OfficeXmlNode fill = data.CreateElement(D, "blipFill");
            fill.Append(fill.CreateElement(D, "blip")).SetAttribute(OfficeXmlNamespaces.Relationships, "embed", relationshipId);
            fill.Append(fill.CreateElement(D, "stretch")).Append(fill.CreateElement(D, "fillRect"));
            SetBackgroundFill(fill, relationshipId);
        }

        /// <summary>Removes the local background so the inherited one shows.</summary>
        public void ClearBackground() {
            ThrowIfRemoved();
            OfficeXmlNode data = CommonSlideData;
            string? imageRelationship = BackgroundImageRelationship(data);
            data.RemoveChildren(P, "bg");
            MarkEdited();
            DeleteImageIfUnused(imageRelationship);
        }

        /// <summary>The shape tree, with the group properties PowerPoint requires added when missing.</summary>
        internal OfficeXmlNode ShapeTree {
            get {
                OfficeXmlNode data = CommonSlideData;
                OfficeXmlNode tree = data.Child(P, "spTree");
                if (tree.IsNone) {
                    tree = data.AddOrdered(P, "spTree", PCommonSlideDataView.Order);
                    MarkEdited();
                }
                if (tree.Child(P, "nvGrpSpPr").IsNone) {
                    OfficeXmlNode group = tree.CreateElement(P, "nvGrpSpPr");
                    AppendDrawingProperties(group, 1U, string.Empty);
                    group.Append(group.CreateElement(P, "cNvGrpSpPr"));
                    group.Append(group.CreateElement(P, "nvPr"));
                    tree.InsertBefore(group, tree.FirstChild);
                    MarkEdited();
                }

                if (tree.Child(P, "grpSpPr").IsNone) {
                    OfficeXmlNode properties = tree.CreateElement(P, "grpSpPr");
                    OfficeXmlNode transform = properties.Append(properties.CreateElement(D, "xfrm"));
                    foreach ((string point, string first, string second) in new[] { ("off", "x", "y"), ("ext", "cx", "cy"), ("chOff", "x", "y"), ("chExt", "cx", "cy") }) {
                        OfficeXmlNode added = transform.Append(transform.CreateElement(D, point));
                        added.SetAttribute(first, "0");
                        added.SetAttribute(second, "0");
                    }
                    tree.Child(P, "nvGrpSpPr").InsertAfterSelf(properties);
                    MarkEdited();
                }

                return tree;
            }
        }

        internal void MarkRemoved() => _removed = true;

        internal void ThrowIfRemoved() {
            if (_removed) throw new InvalidOperationException("This slide master or layout has been removed from the presentation.");
        }

        private static string? BackgroundImageRelationship(OfficeXmlNode data) =>
            data.Child(P, "bg").Child(P, "bgPr").Child(D, "blipFill").Child(D, "blip").Attribute(OfficeXmlNamespaces.Relationships, "embed");

        // Replaces the background with one fill and an empty effect list.
        private void SetBackgroundFill(OfficeXmlNode fill, string? newImage) {
            OfficeXmlNode data = CommonSlideData;
            string? previousImage = BackgroundImageRelationship(data);
            data.RemoveChildren(P, "bg");
            OfficeXmlNode background = data.CreateElement(P, "bg");
            OfficeXmlNode properties = background.Append(background.CreateElement(P, "bgPr"));
            properties.Append(fill);
            properties.Append(properties.CreateElement(D, "effectLst"));
            data.InsertOrdered(background, PCommonSlideDataView.Order);
            MarkEdited();
            if (newImage == null || newImage != previousImage) DeleteImageIfUnused(previousImage);
        }

        private string EmbedImage(Stream image, OfficeImageFormat imageType) {
            ArgumentNullException.ThrowIfNull(image);
            if (!image.CanRead) throw new ArgumentException("Image stream must be readable.", nameof(image));
            using MemoryStream? normalized = PowerPointImageNormalization.NormalizeHeicForEmbedding(image, imageType);
            Stream source = normalized ?? image;
            OfficeImageFormat format = normalized != null ? OfficeImageFormat.Png : imageType;
            if (source.CanSeek) source.Position = 0;
            using var buffer = new MemoryStream();
            source.CopyTo(buffer);
            byte[] data = buffer.ToArray();
            PowerPointPartRef part = PartRef;
            if (!part.IsSdk && Document.TryAddPackageImagePart(part.Name, data, PowerPointImageContentTypes.For(format),
                    PowerPointPartFactory.GetImageExtension(format), out string relationshipId, out _, out _)) {
                return relationshipId;
            }
            return Document.AddSdkSurfaceImage(this, data, format);
        }

        private void DeleteImageIfUnused(string? relationshipId) {
            if (string.IsNullOrEmpty(relationshipId)) return;
            PartRef.RemoveRelationshipIfUnused(relationshipId!, default);
        }

        // A new p:sp with its non-visual properties (id, name, drawing properties, application properties); the
        // application properties come back to receive a placeholder.
        private OfficeXmlNode CreateShape(OfficeXmlNode context, string name, out OfficeXmlNode application) {
            OfficeXmlNode shape = context.CreateElement(P, "sp");
            OfficeXmlNode nonVisual = shape.Append(shape.CreateElement(P, "nvSpPr"));
            AppendDrawingProperties(nonVisual, NextShapeId(), name);
            nonVisual.Append(nonVisual.CreateElement(P, "cNvSpPr"));
            application = nonVisual.Append(nonVisual.CreateElement(P, "nvPr"));
            return shape;
        }

        private static void AppendDrawingProperties(OfficeXmlNode parent, uint id, string name) {
            OfficeXmlNode drawing = parent.Append(parent.CreateElement(P, "cNvPr"));
            drawing.SetUInt32Attribute("id", id);
            drawing.SetAttribute("name", name);
        }

        private uint NextShapeId() {
            uint max = 1U;
            bool any = false;
            foreach (OfficeXmlNode drawing in ShapeTree.Descendants(P, "cNvPr")) {
                uint id = drawing.UInt32Attribute("id") ?? 0U;
                max = any ? Math.Max(max, id) : id;
                any = true;
            }
            return (any ? max : 1U) + 1U;
        }

        private string UniqueName(string baseName) {
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (OfficeXmlNode drawing in ShapeTree.Descendants(P, "cNvPr")) names.Add(drawing.Attribute("name") ?? string.Empty);
            int index = 1;
            string name;
            do {
                name = baseName + " " + index++;
            } while (names.Contains(name));
            return name;
        }

        private uint NextPlaceholderIndex() {
            uint max = 0U;
            foreach (OfficeXmlNode placeholder in ShapeTree.Descendants(P, "ph")) max = Math.Max(max, placeholder.UInt32Attribute("idx") ?? 0U);
            return max + 1U;
        }

        private static bool IsTitle(PowerPointPlaceholderType type) =>
            type is PowerPointPlaceholderType.Title or PowerPointPlaceholderType.CenteredTitle;

        private static bool IsSingleton(PowerPointPlaceholderType type) =>
            IsTitle(type) || type is PowerPointPlaceholderType.SubTitle or PowerPointPlaceholderType.Footer
                or PowerPointPlaceholderType.SlideNumber or PowerPointPlaceholderType.DateAndTime;

        private static OfficeXmlNode CreatePlaceholderParagraph(OfficeXmlNode context, PowerPointPlaceholderType type) {
            string? fieldType = type switch {
                PowerPointPlaceholderType.SlideNumber => "slidenum",
                PowerPointPlaceholderType.DateAndTime => "datetime1",
                _ => null
            };
            OfficeXmlNode paragraph = context.CreateElement(D, "p");
            if (fieldType != null) {
                OfficeXmlNode field = paragraph.Append(paragraph.CreateElement(D, "fld"));
                field.SetAttribute("id", "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}");
                field.SetAttribute("type", fieldType);
                field.Append(field.CreateElement(D, "rPr")).SetAttribute("lang", "en-US");
                field.Append(field.CreateElement(D, "t")).SetInnerText(fieldType == "slidenum" ? "‹#›" : string.Empty);
            }
            paragraph.Append(paragraph.CreateElement(D, "endParaRPr")).SetAttribute("lang", "en-US");
            return paragraph;
        }

        private static OfficeXmlNode CreateTransform(OfficeXmlNode context, OfficeLength left, OfficeLength top, OfficeLength width, OfficeLength height) {
            OfficeXmlNode transform = context.CreateElement(D, "xfrm");
            OfficeXmlNode offset = transform.Append(transform.CreateElement(D, "off"));
            offset.SetInt64Attribute("x", left.ToEmus());
            offset.SetInt64Attribute("y", top.ToEmus());
            OfficeXmlNode extents = transform.Append(transform.CreateElement(D, "ext"));
            extents.SetInt64Attribute("cx", width.ToEmus());
            extents.SetInt64Attribute("cy", height.ToEmus());
            return transform;
        }

        private static OfficeXmlNode CreatePresetGeometry(OfficeXmlNode context, string preset) {
            OfficeXmlNode geometry = context.CreateElement(D, "prstGeom");
            geometry.SetAttribute("prst", preset);
            geometry.Append(geometry.CreateElement(D, "avLst"));
            return geometry;
        }
    }
}
