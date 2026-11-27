using OfficeIMO.Core.Internal.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a grouped set of shapes.
    /// </summary>
    public class PowerPointGroupShape : PowerPointShape {
        internal PowerPointGroupShape(GroupShape groupShape, OpenXmlPartContainer? ownerPart = null) : base(groupShape) {
            _ownerPart = ownerPart;
        }

        private readonly OpenXmlPartContainer? _ownerPart;

        internal PowerPointGroupShape(PowerPointSlide slide, OfficeXmlTree tree, int node) : base(slide, tree, node) {
        }

        internal PowerPointGroupShape(OpenXml.PowerPointPartRef part, OfficeXmlTree tree, int node) : base(part, tree, node) {
        }

        internal GroupShape GroupShape => (GroupShape)Element;

        /// <summary>The slide, layout, or master part whose relationships the group's children use.</summary>
        internal OpenXmlPartContainer? OwnerPart => _ownerPart ?? RelationshipOwnerPart;
    }
}
