using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a grouped set of shapes.
    /// </summary>
    public class PowerPointGroupShape : PowerPointShape {
        internal PowerPointGroupShape(GroupShape groupShape, OpenXmlPartContainer? ownerPart = null) : base(groupShape) {
            OwnerPart = ownerPart;
        }

        internal GroupShape GroupShape => (GroupShape)Element;

        internal OpenXmlPartContainer? OwnerPart { get; }
    }
}
