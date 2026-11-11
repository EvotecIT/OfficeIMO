using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Shared editing surface of a slide master and a slide layout. Shapes, placeholders and the background
    ///     placed here are inherited by every slide that uses the master or layout.
    /// </summary>
    public abstract class PowerPointMasterSurface {
        private bool _removed;

        internal PowerPointMasterSurface(PowerPointDocument document, OpenXmlPart part) {
            Document = document;
            Part = part;
        }

        internal PowerPointDocument Document { get; }

        internal OpenXmlPart Part { get; }

        internal abstract CommonSlideData CommonSlideData { get; }

        /// <summary>Display name shown in PowerPoint's Slide Master view.</summary>
        public string Name {
            get {
                ThrowIfRemoved();
                return CommonSlideData.Name?.Value ?? string.Empty;
            }
            set {
                ThrowIfRemoved();
                ArgumentNullException.ThrowIfNull(value);
                CommonSlideData.Name = value;
            }
        }

        /// <summary>Shapes drawn on this master or layout, in drawing order.</summary>
        public IReadOnlyList<PowerPointShape> Shapes {
            get {
                ThrowIfRemoved();
                var shapes = new List<PowerPointShape>();
                foreach (OpenXmlElement element in ShapeTree.ChildElements) {
                    PowerPointShape? shape = PowerPointSlide.CreateShapeFromElement(element, Part);
                    if (shape != null) shapes.Add(shape);
                }

                return shapes;
            }
        }

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

            uint? resolvedIndex = IsTitle(placeholderType) ? null : index ?? NextPlaceholderIndex();
            var placeholder = new PlaceholderShape { Type = placeholderType.ToOpenXml() };
            if (resolvedIndex.HasValue) placeholder.Index = resolvedIndex.Value;

            Shape shape = new(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = NextShapeId(), Name = name ?? UniqueName(placeholderType + " Placeholder") },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(placeholder)),
                new ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = bounds.Left.ToEmus(), Y = bounds.Top.ToEmus() },
                        new A.Extents { Cx = bounds.Width.ToEmus(), Cy = bounds.Height.ToEmus() }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
                new TextBody(new A.BodyProperties(), new A.ListStyle(), CreatePlaceholderParagraph(placeholderType)));
            ShapeTree.AppendChild(shape);
            return new PowerPointTextBox(shape, Part);
        }

        /// <summary>Adds an auto shape, such as a band or rule, drawn behind slide content on every slide.</summary>
        public PowerPointAutoShape AddShape(OfficePresetShapeType shapeType, OfficeLength left, OfficeLength top,
            OfficeLength width, OfficeLength height, string? name = null) {
            ThrowIfRemoved();
            Shape shape = new(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = NextShapeId(), Name = name ?? UniqueName(shapeType.ToString()) },
                    new NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    CreateTransform(left, top, width, height),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = shapeType.ToOpenXml() }));
            ShapeTree.AppendChild(shape);
            return new PowerPointAutoShape(shape, Part);
        }

        /// <summary>Adds fixed text, such as a confidentiality notice, shown on every slide.</summary>
        public PowerPointTextBox AddTextBox(string text, OfficeLength left, OfficeLength top, OfficeLength width,
            OfficeLength height) {
            ThrowIfRemoved();
            ArgumentNullException.ThrowIfNull(text);
            Shape shape = new(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = NextShapeId(), Name = UniqueName("TextBox") },
                    new NonVisualShapeDrawingProperties { TextBox = true },
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    CreateTransform(left, top, width, height),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                    new A.NoFill()),
                new TextBody(new A.BodyProperties { Wrap = A.TextWrappingValues.Square }, new A.ListStyle(),
                    new A.Paragraph(PowerPointTextContent.CreateInlineContent(text))));
            ShapeTree.AppendChild(shape);
            return new PowerPointTextBox(shape, Part);
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
            Picture picture = new(
                new NonVisualPictureProperties(
                    new NonVisualDrawingProperties { Id = NextShapeId(), Name = UniqueName("Picture") },
                    new NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new BlipFill(new A.Blip { Embed = relationshipId }, new A.Stretch(new A.FillRectangle())),
                new ShapeProperties(
                    CreateTransform(left, top, width, height),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
            ShapeTree.AppendChild(picture);
            var result = new PowerPointImage(picture, Part);
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
                return PowerPointDrawingColor.Read(CommonSlideData.Background?.BackgroundProperties?.GetFirstChild<A.SolidFill>());
            }
            set {
                ThrowIfRemoved();
                if (value == null) {
                    ClearBackground();
                    return;
                }

                SetBackgroundFill(PowerPointDrawingColor.Fill(value.Value, PowerPointDrawingColor.ThemeFor(Part)));
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
            SetBackgroundFill(new A.BlipFill(new A.Blip { Embed = relationshipId }, new A.Stretch(new A.FillRectangle())));
        }

        /// <summary>Removes the local background so the inherited one shows.</summary>
        public void ClearBackground() {
            ThrowIfRemoved();
            string? imageRelationship = CommonSlideData.Background?.BackgroundProperties?
                .GetFirstChild<A.BlipFill>()?.Blip?.Embed?.Value;
            CommonSlideData.Background = null;
            DeleteImageIfUnused(imageRelationship);
        }

        internal ShapeTree ShapeTree {
            get {
                CommonSlideData common = CommonSlideData;
                ShapeTree tree = common.ShapeTree ??= new ShapeTree();
                if (tree.GetFirstChild<NonVisualGroupShapeProperties>() == null) {
                    tree.PrependChild(new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()));
                }

                if (tree.GetFirstChild<GroupShapeProperties>() == null) {
                    tree.InsertAfter(PowerPointUtils.CreateDefaultGroupShapeProperties(),
                        tree.GetFirstChild<NonVisualGroupShapeProperties>());
                }

                return tree;
            }
        }

        internal void MarkRemoved() => _removed = true;

        internal void ThrowIfRemoved() {
            if (_removed) throw new InvalidOperationException("This slide master or layout has been removed from the presentation.");
        }

        private void SetBackgroundFill(OpenXmlElement fill) {
            string? previousImage = CommonSlideData.Background?.BackgroundProperties?
                .GetFirstChild<A.BlipFill>()?.Blip?.Embed?.Value;
            CommonSlideData.Background = new Background(new BackgroundProperties(fill, new A.EffectList()));
            if (fill is not A.BlipFill blip || blip.Blip?.Embed?.Value != previousImage) {
                DeleteImageIfUnused(previousImage);
            }
        }

        private string EmbedImage(Stream image, OfficeImageFormat imageType) {
            ArgumentNullException.ThrowIfNull(image);
            if (!image.CanRead) throw new ArgumentException("Image stream must be readable.", nameof(image));
            using MemoryStream? normalized = PowerPointImageNormalization.NormalizeHeicForEmbedding(image, imageType);
            Stream source = normalized ?? image;
            OfficeImageFormat format = normalized != null ? OfficeImageFormat.Png : imageType;
            string partUri = PowerPointPartFactory.GetIndexedPartUri(Part.OpenXmlPackage, "ppt/media", "image",
                PowerPointPartFactory.GetImageExtension(format), allowBaseWithoutIndex: false);
            ImagePart imagePart = PowerPointPartFactory.CreatePart<ImagePart>(Part, format.ToPartTypeInfo().ContentType, partUri);
            if (source.CanSeek) source.Position = 0;
            imagePart.FeedData(source);
            return Part.GetIdOfPart(imagePart);
        }

        private void DeleteImageIfUnused(string? relationshipId) {
            if (string.IsNullOrEmpty(relationshipId)) return;
            bool used = Part.RootElement?.Descendants<A.Blip>().Any(blip => blip.Embed?.Value == relationshipId) == true;
            if (!used && Part.TryGetPartById(relationshipId!, out OpenXmlPart? part) && part is ImagePart) {
                Part.DeletePart(relationshipId!);
            }
        }

        private uint NextShapeId() =>
            ShapeTree.Descendants<NonVisualDrawingProperties>().Select(p => p.Id?.Value ?? 0U).DefaultIfEmpty(1U).Max() + 1U;

        private string UniqueName(string baseName) {
            var names = new HashSet<string>(ShapeTree.Descendants<NonVisualDrawingProperties>()
                .Select(p => p.Name?.Value ?? string.Empty), StringComparer.Ordinal);
            int index = 1;
            string name;
            do {
                name = baseName + " " + index++;
            } while (names.Contains(name));
            return name;
        }

        private uint NextPlaceholderIndex() {
            uint max = ShapeTree.Descendants<PlaceholderShape>().Select(p => p.Index?.Value ?? 0U).DefaultIfEmpty(0U).Max();
            return max + 1U;
        }

        private static bool IsTitle(PowerPointPlaceholderType type) =>
            type is PowerPointPlaceholderType.Title or PowerPointPlaceholderType.CenteredTitle;

        private static bool IsSingleton(PowerPointPlaceholderType type) =>
            IsTitle(type) || type is PowerPointPlaceholderType.SubTitle or PowerPointPlaceholderType.Footer
                or PowerPointPlaceholderType.SlideNumber or PowerPointPlaceholderType.DateAndTime;

        private static A.Paragraph CreatePlaceholderParagraph(PowerPointPlaceholderType type) {
            string? fieldType = type switch {
                PowerPointPlaceholderType.SlideNumber => "slidenum",
                PowerPointPlaceholderType.DateAndTime => "datetime1",
                _ => null
            };
            if (fieldType == null) return new A.Paragraph(new A.EndParagraphRunProperties { Language = "en-US" });

            var field = new A.Field(new A.RunProperties { Language = "en-US" },
                new A.Text(fieldType == "slidenum" ? "‹#›" : string.Empty)) {
                Id = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}",
                Type = fieldType
            };
            return new A.Paragraph(field, new A.EndParagraphRunProperties { Language = "en-US" });
        }

        private static A.Transform2D CreateTransform(OfficeLength left, OfficeLength top, OfficeLength width, OfficeLength height) =>
            new(new A.Offset { X = left.ToEmus(), Y = top.ToEmus() },
                new A.Extents { Cx = width.ToEmus(), Cy = height.ToEmus() });
    }
}
