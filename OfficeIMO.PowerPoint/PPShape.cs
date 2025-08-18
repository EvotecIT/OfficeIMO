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
    }
}

