using OfficeIMO.Core.Internal.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Base class for shapes used on PowerPoint slides.
    /// </summary>
    public abstract partial class PowerPointShape {
        private const int EmusPerPoint = 12700;

        private OpenXmlElement? _element;
        private OpenXmlPart? _relationshipOwnerPart;

        internal PowerPointShape(OpenXmlElement element, OpenXmlPartContainer? ownerPart = null) {
            _element = element;
            RelationshipOwnerPart = ownerPart as OpenXmlPart;
        }

        // A shape of a slide whose SDK part is not open: it reads and edits its element in the slide's package tree, and
        // anything it does not do there opens the SDK and binds the shape to the same element there.
        internal PowerPointShape(PowerPointSlide slide, OfficeXmlTree tree, int node) {
            OwnerSlide = slide;
            PackageTree = tree;
            PackageNode = node;
        }

        internal OpenXmlElement Element => _element ?? BindSdkElement();
        internal PowerPointSlide? OwnerSlide { get; private set; }

        internal OpenXmlPart? RelationshipOwnerPart {
            get => _relationshipOwnerPart ?? (PackageTree != null ? OwnerSlide?.SlidePart : null);
            private set => _relationshipOwnerPart = value;
        }

        /// <summary>The slide tree that holds the shape while it works without the SDK, and its element there.</summary>
        internal OfficeXmlTree? PackageTree { get; }
        internal int PackageNode { get; }

        /// <summary>Whether the shape works on its slide's package tree rather than an SDK element.</summary>
        internal bool IsPackageBacked => _element == null;

        // The SDK reads the bytes the tree wrote, so the element sits at the same place: the same element child at each
        // level down from the slide root. Comments and text the SDK keeps as nodes are not elements of the tree.
        private OpenXmlElement BindSdkElement() {
            OfficeXmlTree tree = PackageTree!;
            var path = new List<int>();
            for (int node = PackageNode; node != tree.Root; node = tree.ParentOf(node)) {
                int index = 0;
                for (int sibling = tree.PreviousSiblingOf(node); sibling != OfficeXmlTree.NoNode; sibling = tree.PreviousSiblingOf(sibling)) {
                    if (tree.KindOf(sibling) == OfficeXmlTreeNodeKind.Element) index++;
                }
                path.Add(index);
            }
            string localName = tree.LocalNameOf(PackageNode);
            OpenXmlElement element = OwnerSlide!.SlidePart.Slide
                ?? throw new InvalidOperationException("The slide the SDK opened has no content.");
            for (int level = path.Count - 1; level >= 0; level--) {
                OpenXmlElement? child = null;
                int remaining = path[level];
                foreach (OpenXmlElement candidate in element.ChildElements) {
                    if (candidate is OpenXmlMiscNode) continue;
                    if (remaining-- == 0) {
                        child = candidate;
                        break;
                    }
                }
                element = child ?? throw new InvalidOperationException("The shape is not where the slide the SDK opened has it.");
            }
            if (!string.Equals(element.LocalName, localName, StringComparison.Ordinal)) {
                throw new InvalidOperationException("The shape is not where the slide the SDK opened has it.");
            }
            _element = element;
            return element;
        }

        internal PowerPointShape AttachTo(PowerPointSlide slide, OpenXmlPartContainer? ownerPart = null) {
            OwnerSlide = slide;
            RelationshipOwnerPart = ownerPart as OpenXmlPart ?? RelationshipOwnerPart ?? slide.SlidePart;
            return this;
        }

        /// <summary>
        ///     Numeric shape identifier stored in the PowerPoint non-visual drawing properties.
        /// </summary>
        public uint? Id {
            get => GetNonVisualDrawingProperties(create: false)?.Id?.Value;
        }

        /// <summary>
        ///     Name assigned to the shape.
        /// </summary>
        public string? Name {
            get => GetNonVisualDrawingProperties(create: false)?.Name?.Value;
            set {
                NonVisualDrawingProperties drawing = GetNonVisualDrawingProperties(create: true)
                    ?? throw new NotSupportedException("This shape type does not expose non-visual drawing properties.");
                drawing.Name = value ?? string.Empty;
            }
        }

        /// <summary>
        ///     Alternative text description assigned to the shape.
        /// </summary>
        public string? AltText {
            get => GetNonVisualDrawingProperties(create: false)?.Description?.Value;
            set {
                NonVisualDrawingProperties drawing = GetNonVisualDrawingProperties(create: true)
                    ?? throw new NotSupportedException("This shape type does not expose non-visual drawing properties.");
                drawing.Description = string.IsNullOrWhiteSpace(value)
                    ? null
                    : new DocumentFormat.OpenXml.StringValue(value);
            }
        }

        /// <summary>
        ///     Gets or sets whether the shape is hidden.
        /// </summary>
        public bool Hidden {
            get => GetNonVisualDrawingProperties(create: false)?.Hidden?.Value == true;
            set {
                NonVisualDrawingProperties drawing = GetNonVisualDrawingProperties(create: true)
                    ?? throw new NotSupportedException("This shape type does not expose non-visual drawing properties.");
                drawing.Hidden = value ? true : null;
            }
        }

        /// <summary>
        ///     Click hyperlink assigned to the shape, when present. Internal slide links are
        ///     returned as stable Markdown-compatible fragments such as <c>#slide-2</c>.
        /// </summary>
        public Uri? Hyperlink {
            get {
                if (OwnerSlide == null) return null;
                return PowerPointHyperlinkResolver.Resolve(
                    RelationshipOwnerPart ?? OwnerSlide.SlidePart, OwnerSlide.SlidePart,
                    GetNonVisualDrawingProperties(create: false)?
                        .GetFirstChild<A.HyperlinkOnClick>());
            }
            set {
                if (value == null) ClearHyperlink();
                else SetHyperlink(value);
            }
        }

        /// <summary>
        ///     Placeholder type associated with the shape, if any.
        /// </summary>
        public PowerPointPlaceholderType? ShapePlaceholderType => GetPlaceholderShape()?.Type?.Value.ToOfficeEnum();

        /// <summary>
        ///     Placeholder index associated with the shape, if any.
        /// </summary>
        public uint? ShapePlaceholderIndex => GetPlaceholderShape()?.Index?.Value;

        /// <summary>
        ///     Preferred placeholder size associated with the shape, if any.
        /// </summary>
        public PowerPointPlaceholderSize? ShapePlaceholderSize => GetPlaceholderShape()?.Size?.Value.ToOfficeEnum();

        /// <summary>
        ///     Placeholder orientation associated with the shape, if any.
        /// </summary>
        public PowerPointPlaceholderDirection? ShapePlaceholderOrientation =>
            GetPlaceholderShape()?.Orientation?.Value.ToOfficeEnum();

        /// <summary>
        ///     Primary content type represented by this shape wrapper.
        /// </summary>
        public PowerPointShapeContentType ShapeContentType => Element switch {
            GroupShape => PowerPointShapeContentType.Group,
            Picture p when IsMediaPicture(p) => PowerPointShapeContentType.Media,
            Picture => PowerPointShapeContentType.Picture,
            GraphicFrame g when g.Graphic?.GraphicData?.GetFirstChild<A.Table>() != null => PowerPointShapeContentType.Table,
            GraphicFrame g when g.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>() != null => PowerPointShapeContentType.Chart,
            GraphicFrame g when g.Graphic?.GraphicData?.GetFirstChild<Dgm.RelationshipIds>() != null => PowerPointShapeContentType.SmartArt,
            GraphicFrame g when g.Graphic?.GraphicData?.GetFirstChild<OleObject>() != null => PowerPointShapeContentType.OleObject,
            ConnectionShape => PowerPointShapeContentType.Connector,
            Shape s when s.TextBody != null => PowerPointShapeContentType.TextBox,
            Shape => PowerPointShapeContentType.AutoShape,
            _ => PowerPointShapeContentType.Unknown
        };

        /// <summary>
        ///     Zero-based drawing order within the parent shape tree, where larger values render above earlier shapes.
        /// </summary>
        public int DrawingOrder {
            get {
                OpenXmlElement? parent = Element.Parent;
                if (parent == null) {
                    return -1;
                }

                int order = 0;
                foreach (OpenXmlElement child in parent.ChildElements) {
                    if (!IsDrawingElement(child)) {
                        continue;
                    }

                    if (ReferenceEquals(child, Element)) {
                        return order;
                    }

                    order++;
                }

                return -1;
            }
        }

        /// <summary>
        ///     Removes this shape from its owning slide or parent shape tree.
        /// </summary>
        public void Remove() {
            if (OwnerSlide != null) {
                OwnerSlide.RemoveShape(this);
                return;
            }

            Element.Remove();
        }

        /// <summary>
        ///     Duplicates this shape on its owning slide.
        /// </summary>
        public PowerPointShape Duplicate(OfficeLength offsetX = default, OfficeLength offsetY = default) {
            if (OwnerSlide == null) {
                throw new InvalidOperationException("Shape duplication requires a shape attached to a slide.");
            }

            return OwnerSlide.DuplicateShape(this, offsetX, offsetY);
        }




        private static bool IsDrawingElement(OpenXmlElement element) =>
            element is Shape or ConnectionShape or Picture or GraphicFrame or GroupShape;
    }
}
