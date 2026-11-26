using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>The SDK constructor of <see cref="PowerPointAutoShape"/>, for shapes read or added through the SDK.</summary>
    public partial class PowerPointAutoShape {
        internal PowerPointAutoShape(Shape shape, OpenXmlPartContainer? ownerPart = null) : base(shape, ownerPart) {
        }
    }
}
