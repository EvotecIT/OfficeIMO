using DocumentFormat.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Base class for shapes used on PowerPoint slides.
    /// </summary>
    public abstract class PPShape {
        internal OpenXmlElement Element { get; }

        internal PPShape(OpenXmlElement element) {
            Element = element;
        }
    }
}

