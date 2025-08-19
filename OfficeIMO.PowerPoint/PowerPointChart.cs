using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a chart on a slide.
    /// </summary>
    public class PowerPointChart : PowerPointShape {
        internal PowerPointChart(GraphicFrame frame) : base(frame) {
        }
    }
}