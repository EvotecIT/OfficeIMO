using DocumentFormat.OpenXml;
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

        /// <summary>Binds a package-backed shape to its SDK element now, while the slide the SDK opened is unchanged.</summary>
        internal void BindSdk() {
            if (_element != null || PackageTree == null) return;
            try {
                BindSdkElement();
            } catch (InvalidOperationException) {
                // A shape the opened slide doesn't have where expected reports it when it is used, not while the SDK opens.
            }
        }
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

        private OpenXmlElement BindSdkElement() {
            OpenXmlElement element = PowerPointSlide.BindPackageElement(OwnerSlide!, PackageTree!, PackageNode);
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
            get => TryGetPackageId(out uint? id) ? id : GetNonVisualDrawingProperties(create: false)?.Id?.Value;
        }

        /// <summary>
        ///     Name assigned to the shape.
        /// </summary>
        public string? Name {
            get => TryReadPackageDrawingAttribute("name", out string? name) ? name : GetNonVisualDrawingProperties(create: false)?.Name?.Value;
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
            get => TryReadPackageDrawingAttribute("descr", out string? description) ? description : GetNonVisualDrawingProperties(create: false)?.Description?.Value;
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
            get => TryGetPackageHidden(out bool hidden) ? hidden : GetNonVisualDrawingProperties(create: false)?.Hidden?.Value == true;
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
