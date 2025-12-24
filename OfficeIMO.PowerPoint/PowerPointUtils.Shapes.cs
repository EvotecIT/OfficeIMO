using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointUtils {
        internal static GroupShapeProperties CreateDefaultGroupShapeProperties() {
            return new GroupShapeProperties(CreateDefaultTransformGroup());
        }

        private static D.TransformGroup CreateDefaultTransformGroup() {
            return new D.TransformGroup(
                new D.Offset { X = 0L, Y = 0L },
                new D.Extents { Cx = 0L, Cy = 0L },
                new D.ChildOffset { X = 0L, Y = 0L },
                new D.ChildExtents { Cx = 0L, Cy = 0L });
        }

    }
}
