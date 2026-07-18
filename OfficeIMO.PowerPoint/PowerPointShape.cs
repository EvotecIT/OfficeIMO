using System;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Base class for shapes used on PowerPoint slides.
    /// </summary>
    public abstract partial class PowerPointShape {
        private const int EmusPerPoint = 12700;

        internal PowerPointShape(OpenXmlElement element) {
            Element = element;
        }

        internal OpenXmlElement Element { get; }
        internal PowerPointSlide? OwnerSlide { get; private set; }

        internal PowerPointShape AttachTo(PowerPointSlide slide) {
            OwnerSlide = slide;
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
                    OwnerSlide.SlidePart, OwnerSlide.SlidePart,
                    GetNonVisualDrawingProperties(create: false)?
                        .GetFirstChild<A.HyperlinkOnClick>());
            }
        }

        /// <summary>
        ///     Placeholder type associated with the shape, if any.
        /// </summary>
        public PlaceholderValues? ShapePlaceholderType => GetPlaceholderShape()?.Type?.Value;

        /// <summary>
        ///     Placeholder index associated with the shape, if any.
        /// </summary>
        public uint? ShapePlaceholderIndex => GetPlaceholderShape()?.Index?.Value;

        /// <summary>
        ///     Preferred placeholder size associated with the shape, if any.
        /// </summary>
        public PlaceholderSizeValues? ShapePlaceholderSize => GetPlaceholderShape()?.Size?.Value;

        /// <summary>
        ///     Placeholder orientation associated with the shape, if any.
        /// </summary>
        public DirectionValues? ShapePlaceholderOrientation =>
            GetPlaceholderShape()?.Orientation?.Value;

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
        public PowerPointShape Duplicate(long offsetX = 0L, long offsetY = 0L) {
            if (OwnerSlide == null) {
                throw new InvalidOperationException("Shape duplication requires a shape attached to a slide.");
            }

            return OwnerSlide.DuplicateShape(this, offsetX, offsetY);
        }

        /// <summary>
        ///     Duplicates this shape on its owning slide and offsets it in centimeters.
        /// </summary>
        public PowerPointShape DuplicateCm(double offsetXCm, double offsetYCm) {
            return Duplicate(PowerPointUnits.FromCentimeters(offsetXCm), PowerPointUnits.FromCentimeters(offsetYCm));
        }

        /// <summary>
        ///     Duplicates this shape on its owning slide and offsets it in inches.
        /// </summary>
        public PowerPointShape DuplicateInches(double offsetXInches, double offsetYInches) {
            return Duplicate(PowerPointUnits.FromInches(offsetXInches), PowerPointUnits.FromInches(offsetYInches));
        }

        /// <summary>
        ///     Duplicates this shape on its owning slide and offsets it in points.
        /// </summary>
        public PowerPointShape DuplicatePoints(double offsetXPoints, double offsetYPoints) {
            return Duplicate(PowerPointUnits.FromPoints(offsetXPoints), PowerPointUnits.FromPoints(offsetYPoints));
        }

        private static bool IsDrawingElement(OpenXmlElement element) =>
            element is Shape or ConnectionShape or Picture or GraphicFrame or GroupShape;
    }
}
