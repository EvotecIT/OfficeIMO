using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

internal static partial class PowerPointSlideImageRenderer {
    private static bool HasShapeFillGradient(PowerPointShape source) =>
        TryGetShapeFillGradient(source, out _, out _);
}
