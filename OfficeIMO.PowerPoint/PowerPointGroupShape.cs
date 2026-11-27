using OfficeIMO.Core.Internal.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a grouped set of shapes.
    /// </summary>
    public class PowerPointGroupShape : PowerPointShape {
        private readonly OpenXmlPartContainer? _ownerPart;
        private readonly bool _readFromPackage;

        internal PowerPointGroupShape(GroupShape groupShape, OpenXmlPartContainer? ownerPart = null) : base(groupShape) {
            _ownerPart = ownerPart;
        }

        internal PowerPointGroupShape(PowerPointSlide slide, OfficeXmlTree tree, int node) : base(slide, tree, node) {
            _readFromPackage = true;
        }

        internal GroupShape GroupShape => (GroupShape)Element;

        // A group read from the package belongs to its slide's part once the SDK opens.
        internal OpenXmlPartContainer? OwnerPart => _ownerPart ?? (_readFromPackage ? OwnerSlide?.SlidePart : null);
    }
}
