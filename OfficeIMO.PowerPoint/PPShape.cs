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

        private A.Transform2D EnsureTransform2D(ShapeProperties props) {
            props.Transform2D ??= new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 0L, Cy = 0L });
            props.Transform2D.Offset ??= new A.Offset { X = 0L, Y = 0L };
            props.Transform2D.Extents ??= new A.Extents { Cx = 0L, Cy = 0L };
            return props.Transform2D;
        }

        private Transform EnsureTransform(GraphicFrame frame) {
            frame.Transform ??= new Transform(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 0L, Cy = 0L });
            frame.Transform.Offset ??= new A.Offset { X = 0L, Y = 0L };
            frame.Transform.Extents ??= new A.Extents { Cx = 0L, Cy = 0L };
            return frame.Transform;
        }

        /// <summary>
        /// Left position of the shape in English Metric Units (EMU).
        /// </summary>
        public long Left {
            get {
                return Element switch {
                    Shape s => s.ShapeProperties?.Transform2D?.Offset?.X ?? 0L,
                    Picture p => p.ShapeProperties?.Transform2D?.Offset?.X ?? 0L,
                    GraphicFrame g => g.Transform?.Offset?.X ?? 0L,
                    _ => 0L,
                };
            }
            set {
                switch (Element) {
                    case Shape s:
                        A.Transform2D ts = EnsureTransform2D(s.ShapeProperties ??= new ShapeProperties());
                        ts.Offset!.X = value;
                        break;
                    case Picture p:
                        A.Transform2D tp = EnsureTransform2D(p.ShapeProperties ??= new ShapeProperties());
                        tp.Offset!.X = value;
                        break;
                    case GraphicFrame g:
                        Transform tg = EnsureTransform(g);
                        tg.Offset!.X = value;
                        break;
                }
            }
        }

        /// <summary>
        /// Top position of the shape in English Metric Units (EMU).
        /// </summary>
        public long Top {
            get {
                return Element switch {
                    Shape s => s.ShapeProperties?.Transform2D?.Offset?.Y ?? 0L,
                    Picture p => p.ShapeProperties?.Transform2D?.Offset?.Y ?? 0L,
                    GraphicFrame g => g.Transform?.Offset?.Y ?? 0L,
                    _ => 0L,
                };
            }
            set {
                switch (Element) {
                    case Shape s:
                        A.Transform2D ts = EnsureTransform2D(s.ShapeProperties ??= new ShapeProperties());
                        ts.Offset!.Y = value;
                        break;
                    case Picture p:
                        A.Transform2D tp = EnsureTransform2D(p.ShapeProperties ??= new ShapeProperties());
                        tp.Offset!.Y = value;
                        break;
                    case GraphicFrame g:
                        Transform tg = EnsureTransform(g);
                        tg.Offset!.Y = value;
                        break;
                }
            }
        }

        /// <summary>
        /// Width of the shape in English Metric Units (EMU).
        /// </summary>
        public long Width {
            get {
                return Element switch {
                    Shape s => s.ShapeProperties?.Transform2D?.Extents?.Cx ?? 0L,
                    Picture p => p.ShapeProperties?.Transform2D?.Extents?.Cx ?? 0L,
                    GraphicFrame g => g.Transform?.Extents?.Cx ?? 0L,
                    _ => 0L,
                };
            }
            set {
                switch (Element) {
                    case Shape s:
                        A.Transform2D ts = EnsureTransform2D(s.ShapeProperties ??= new ShapeProperties());
                        ts.Extents!.Cx = value;
                        break;
                    case Picture p:
                        A.Transform2D tp = EnsureTransform2D(p.ShapeProperties ??= new ShapeProperties());
                        tp.Extents!.Cx = value;
                        break;
                    case GraphicFrame g:
                        Transform tg = EnsureTransform(g);
                        tg.Extents!.Cx = value;
                        break;
                }
            }
        }

        /// <summary>
        /// Height of the shape in English Metric Units (EMU).
        /// </summary>
        public long Height {
            get {
                return Element switch {
                    Shape s => s.ShapeProperties?.Transform2D?.Extents?.Cy ?? 0L,
                    Picture p => p.ShapeProperties?.Transform2D?.Extents?.Cy ?? 0L,
                    GraphicFrame g => g.Transform?.Extents?.Cy ?? 0L,
                    _ => 0L,
                };
            }
            set {
                switch (Element) {
                    case Shape s:
                        A.Transform2D ts = EnsureTransform2D(s.ShapeProperties ??= new ShapeProperties());
                        ts.Extents!.Cy = value;
                        break;
                    case Picture p:
                        A.Transform2D tp = EnsureTransform2D(p.ShapeProperties ??= new ShapeProperties());
                        tp.Extents!.Cy = value;
                        break;
                    case GraphicFrame g:
                        Transform tg = EnsureTransform(g);
                        tg.Extents!.Cy = value;
                        break;
                }
            }
        }
    }
}

