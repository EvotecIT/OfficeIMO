using System.Collections.Generic;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        // Read only indices in the already bounded cache; sparse or malformed dPt
        // indices must neither allocate by index nor change category alignment.
        private static IReadOnlyList<OfficeColor?>? ReadPointColors(
            OpenXmlCompositeElement series, int pointCount, A.ColorScheme? colorScheme) {
            OfficeColor?[]? colors = null;
            foreach (C.DataPoint point in GetBoundedCachedPoints(series.Elements<C.DataPoint>())) {
                uint? index = point.GetFirstChild<C.Index>()?.Val?.Value;
                if (!index.HasValue || index.Value >= (uint)pointCount) continue;
                A.SolidFill? fill = point.GetFirstChild<C.ChartShapeProperties>()?
                    .GetFirstChild<A.SolidFill>();
                OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(fill, colorScheme);
                if (!color.HasValue) continue;
                colors ??= new OfficeColor?[pointCount];
                colors[(int)index.Value] = color;
            }
            return colors;
        }
    }
}
