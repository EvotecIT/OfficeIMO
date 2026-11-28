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
        private readonly bool _readFromSlidePackage;

        internal PowerPointGroupShape(PowerPointSlide slide, OfficeXmlTree tree, int node) : base(slide, tree, node) {
            _readFromSlidePackage = true;
        }

        internal PowerPointGroupShape(OpenXml.PowerPointPartRef part, OfficeXmlTree tree, int node) : base(part, tree, node) {
        }

        internal GroupShape GroupShape => (GroupShape)Element;

        /// <summary>The part whose relationships the group's children use: a master's or layout's for a group held there.</summary>
        /// <remarks>A group read from a slide's package belongs to that slide's part once the SDK opens.</remarks>
        internal OpenXmlPartContainer? OwnerPart => _ownerPart
            ?? (!PackagePart.IsNone ? RelationshipOwnerPart : _readFromSlidePackage ? OwnerSlide?.SlidePart : null);
    }
}
