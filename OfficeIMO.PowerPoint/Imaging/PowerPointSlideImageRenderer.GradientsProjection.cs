using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint;

internal static partial class PowerPointSlideImageRenderer {
    private static bool HasShapeFillGradient(PowerPointShape source) =>
        TryGetShapeFillGradient(source, out _, out _);
}
