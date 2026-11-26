using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Base class for shapes used on PowerPoint slides.
    /// </summary>
    public abstract partial class PowerPointShape {
        private const int EmusPerPoint = 12700;

        /// <summary>
        ///     Numeric shape identifier stored in the PowerPoint non-visual drawing properties.
        /// </summary>
        public uint? Id => NonVisualDrawingElement(create: false).UInt32Attribute("id");

        /// <summary>
        ///     Name assigned to the shape.
        /// </summary>
        public string? Name {
            get => NonVisualDrawingElement(create: false).Attribute("name");
            set {
                RequireNonVisualDrawing().SetAttribute("name", value ?? string.Empty);
                OwnerPartRef.MarkEdited();
            }
        }

        /// <summary>
        ///     Alternative text description assigned to the shape.
        /// </summary>
        public string? AltText {
            get => NonVisualDrawingElement(create: false).Attribute("descr");
            set {
                RequireNonVisualDrawing().SetAttribute("descr", string.IsNullOrWhiteSpace(value) ? null : value);
                OwnerPartRef.MarkEdited();
            }
        }

        /// <summary>
        ///     Gets or sets whether the shape is hidden.
        /// </summary>
        public bool Hidden {
            get => NonVisualDrawingElement(create: false).BooleanAttribute("hidden") == true;
            set {
                RequireNonVisualDrawing().SetAttribute("hidden", value ? "1" : null);
                OwnerPartRef.MarkEdited();
            }
        }

        private OfficeXmlNode RequireNonVisualDrawing() {
            OfficeXmlNode drawing = NonVisualDrawingElement(create: true);
            return drawing.IsNone ? throw new NotSupportedException("This shape type does not expose non-visual drawing properties.") : drawing;
        }

        /// <summary>
        ///     Click hyperlink assigned to the shape, when present. Internal slide links are
        ///     returned as stable Markdown-compatible fragments such as <c>#slide-2</c>.
        /// </summary>
        public Uri? Hyperlink {
            get {
                if (OwnerSlide == null) return null;
                PowerPointPartRef owner = OwnerPartRef;
                return PowerPointHyperlinkResolver.Resolve(owner.IsNone ? OwnerSlide.PartRef : owner, OwnerSlide.PartRef,
                    NonVisualDrawingElement(create: false).Child(OfficeXmlNamespaces.DrawingML, "hlinkClick"));
            }
            set {
                if (value == null) ClearHyperlink();
                else SetHyperlink(value);
            }
        }

        /// <summary>
        ///     Placeholder type associated with the shape, if any.
        /// </summary>
        public PowerPointPlaceholderType? ShapePlaceholderType =>
            ShapePlaceholderElement.Attribute("type") is { } token && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out PowerPointPlaceholderType type) ? type : null;

        /// <summary>
        ///     Placeholder index associated with the shape, if any.
        /// </summary>
        public uint? ShapePlaceholderIndex => ShapePlaceholderElement.UInt32Attribute("idx");

        /// <summary>
        ///     Preferred placeholder size associated with the shape, if any.
        /// </summary>
        public PowerPointPlaceholderSize? ShapePlaceholderSize =>
            ShapePlaceholderElement.Attribute("sz") is { } token && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out PowerPointPlaceholderSize size) ? size : null;

        /// <summary>
        ///     Placeholder orientation associated with the shape, if any.
        /// </summary>
        public PowerPointPlaceholderDirection? ShapePlaceholderOrientation =>
            ShapePlaceholderElement.Attribute("orient") is { } token && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out PowerPointPlaceholderDirection direction) ? direction : null;

        /// <summary>
        ///     Primary content type represented by this shape wrapper.
        /// </summary>
        public PowerPointShapeContentType ShapeContentType {
            get {
                OfficeXmlNode element = XmlElement;
                if (element.NamespaceId != OfficeXmlNamespaces.PresentationML) return PowerPointShapeContentType.Unknown;
                switch (element.LocalName) {
                    case "grpSp":
                        return PowerPointShapeContentType.Group;
                    case "pic":
                        return IsMediaPicture(element) ? PowerPointShapeContentType.Media : PowerPointShapeContentType.Picture;
                    case "graphicFrame": {
                        OfficeXmlNode data = element.Child(OfficeXmlNamespaces.DrawingML, "graphic").Child(OfficeXmlNamespaces.DrawingML, "graphicData");
                        if (!data.Child(OfficeXmlNamespaces.DrawingML, "tbl").IsNone) return PowerPointShapeContentType.Table;
                        if (!data.Child(OfficeXmlNamespaces.Chart, "chart").IsNone) return PowerPointShapeContentType.Chart;
                        if (!data.Child(OfficeXmlNamespaces.Diagram, "relIds").IsNone) return PowerPointShapeContentType.SmartArt;
                        if (!data.Child(OfficeXmlNamespaces.PresentationML, "oleObj").IsNone) return PowerPointShapeContentType.OleObject;
                        return PowerPointShapeContentType.Unknown;
                    }
                    case "cxnSp":
                        return PowerPointShapeContentType.Connector;
                    case "sp":
                        return element.Child(OfficeXmlNamespaces.PresentationML, "txBody").IsNone ? PowerPointShapeContentType.AutoShape : PowerPointShapeContentType.TextBox;
                    default:
                        return PowerPointShapeContentType.Unknown;
                }
            }
        }

        /// <summary>
        ///     Zero-based drawing order within the parent shape tree, where larger values render above earlier shapes.
        /// </summary>
        public int DrawingOrder {
            get {
                OfficeXmlNode element = XmlElement;
                OfficeXmlNode parent = element.Parent;
                if (parent.IsNone) return -1;
                int order = 0;
                foreach (OfficeXmlNode child in parent.Children()) {
                    if (!IsDrawingElement(child)) continue;
                    if (child.Equals(element)) return order;
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

            XmlElement.Remove();
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




        private static bool IsDrawingElement(OfficeXmlNode element) =>
            element.NamespaceId == OfficeXmlNamespaces.PresentationML && element.LocalName is "sp" or "cxnSp" or "pic" or "graphicFrame" or "grpSp";
    }
}
