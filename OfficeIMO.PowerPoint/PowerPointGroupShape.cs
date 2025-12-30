using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a grouped set of shapes.
    /// </summary>
    public class PowerPointGroupShape : PowerPointShape {
        internal PowerPointGroupShape(GroupShape groupShape) : base(groupShape) {
        }

        internal GroupShape GroupShape => (GroupShape)Element;
    }
}
