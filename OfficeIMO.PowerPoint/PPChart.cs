using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents a chart on a slide.
    /// </summary>
    public class PPChart : PPShape {
        internal PPChart(GraphicFrame frame) : base(frame) {
        }
    }
}
