using System;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Base class for shapes used on PowerPoint slides.
    /// </summary>
    public abstract class PPShape {
        internal OpenXmlElement Element { get; }

        internal PPShape(OpenXmlElement element) {
            Element = element;
        }

        /// <summary>
        /// Name assigned to the shape.
        /// </summary>
        public string? Name {
            get {
                switch (Element) {
                    case Shape s:
                        return s.NonVisualShapeProperties?.NonVisualDrawingProperties.Name?.Value;
                    case Picture p:
                        return p.NonVisualPictureProperties?.NonVisualDrawingProperties.Name?.Value;
                    case GraphicFrame g:
                        return g.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties.Name?.Value;
                    default:
                        return null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the fill color of the shape in hex format (e.g. "FF0000").
        /// </summary>
        public string? FillColor {
            get {
                if (Element is Shape shape) {
                    A.SolidFill? solid = shape.ShapeProperties?.GetFirstChild<A.SolidFill>();
                    return solid?.RgbColorModelHex?.Val;
                }
                return null;
            }
            set {
                if (Element is Shape shape) {
                    shape.ShapeProperties ??= new ShapeProperties();
                    shape.ShapeProperties.RemoveAllChildren<A.SolidFill>();
                    if (value != null) {
                        shape.ShapeProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = value }));
                    }
                }
            }
        }

        private A.Offset GetOffset() {
            switch (Element) {
                case Shape s:
                    s.ShapeProperties ??= new ShapeProperties();
                    s.ShapeProperties.Transform2D ??= new A.Transform2D();
                    s.ShapeProperties.Transform2D.Offset ??= new A.Offset();
                    return s.ShapeProperties.Transform2D.Offset;
                case Picture p:
                    p.ShapeProperties ??= new ShapeProperties();
                    p.ShapeProperties.Transform2D ??= new A.Transform2D();
                    p.ShapeProperties.Transform2D.Offset ??= new A.Offset();
                    return p.ShapeProperties.Transform2D.Offset;
                case GraphicFrame g:
                    g.Transform ??= new Transform();
                    g.Transform.Offset ??= new A.Offset();
                    return g.Transform.Offset;
                default:
                    throw new NotSupportedException();
            }
        }

        private A.Extents GetExtents() {
            switch (Element) {
                case Shape s:
                    s.ShapeProperties ??= new ShapeProperties();
                    s.ShapeProperties.Transform2D ??= new A.Transform2D();
                    s.ShapeProperties.Transform2D.Extents ??= new A.Extents();
                    return s.ShapeProperties.Transform2D.Extents;
                case Picture p:
                    p.ShapeProperties ??= new ShapeProperties();
                    p.ShapeProperties.Transform2D ??= new A.Transform2D();
                    p.ShapeProperties.Transform2D.Extents ??= new A.Extents();
                    return p.ShapeProperties.Transform2D.Extents;
                case GraphicFrame g:
                    g.Transform ??= new Transform();
                    g.Transform.Extents ??= new A.Extents();
                    return g.Transform.Extents;
                default:
                    throw new NotSupportedException();
            }
        }

        /// <summary>
        /// Horizontal position of the shape in EMUs.
        /// </summary>
        public long Left {
            get => GetOffset().X?.Value ?? 0L;
            set => GetOffset().X = value;
        }

        /// <summary>
        /// Vertical position of the shape in EMUs.
        /// </summary>
        public long Top {
            get => GetOffset().Y?.Value ?? 0L;
            set => GetOffset().Y = value;
        }

        /// <summary>
        /// Width of the shape in EMUs.
        /// </summary>
        public long Width {
            get => GetExtents().Cx?.Value ?? 0L;
            set => GetExtents().Cx = value;
        }

        /// <summary>
        /// Height of the shape in EMUs.
        /// </summary>
        public long Height {
            get => GetExtents().Cy?.Value ?? 0L;
            set => GetExtents().Cy = value;
        }
    }
}

