using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

internal static partial class OfficeOpenXmlChartSeriesReader {
    private static List<T> OrderSeries<T>(List<T> elements) where T : OpenXmlCompositeElement {
        // Bounds are established before ordering. Missing order retains legacy XML
        // ordering; explicit orders must be unique to avoid ambiguous projections.
        var orders = new HashSet<uint>();
        foreach (var element in elements) {
            uint? order = element.GetFirstChild<C.Order>()?.Val?.Value;
            if (order.HasValue && !orders.Add(order.Value))
                throw new InvalidDataException("Native chart series have duplicate plotting orders.");
        }
        return elements.OrderBy(element => element.GetFirstChild<C.Order>()?.Val?.Value ?? uint.MaxValue).ToList();
    }

    private static bool IsSupportedSeriesShape(C.ChartShapeProperties? properties, A.ColorScheme? scheme, bool filled, bool area = false, bool connectLine = true) {
        if (properties == null) return true;
        foreach (var child in properties.ChildElements) {
            if ((child is A.EffectList || child is A.Shape3DType) && !child.HasChildren && !child.HasAttributes) continue;
            if (child is A.SolidFill && OfficeOpenXmlThemeColorResolver.ResolveColor(child, scheme).HasValue) continue;
            if (child is A.NoFill && !filled) continue;
            if (child is not A.Outline outline) return false;
            // Marker-only series retain native line metadata that has no rendered
            // effect. Filled series still use the outline around their geometry.
            if (!filled && !connectLine) continue;
            if (outline.Width?.Value == 0 && outline.GetFirstChild<A.NoFill>() == null) return false;
            if (outline.CapType != null || outline.Alignment != null || outline.CompoundLineType != null || outline.Width?.Value < 0) return false;
            foreach (var lineChild in outline.ChildElements) {
                if (lineChild is A.NoFill) continue;
                if (lineChild is A.SolidFill && OfficeOpenXmlThemeColorResolver.ResolveColor(lineChild, scheme).HasValue) continue;
                if ((!filled || area) && lineChild is A.PresetDash && (!connectLine || ReadDash(outline).HasValue)) continue;
                return false;
            }
        }
        return true;
    }

    private static IReadOnlyList<OfficeChartPointStyle?>? InheritFilledOutline(
        IReadOnlyList<OfficeChartPointStyle?>? styles, int count, A.Outline? outline, OfficeColor? color, double? width) {
        if (outline == null || (!outline.HasChildren && !width.HasValue)) return styles;
        bool visible = outline.GetFirstChild<A.NoFill>() == null;
        var result = new OfficeChartPointStyle?[count];
        for (int index = 0; index < count; index++) {
            var point = styles?[index];
            result[index] = new OfficeChartPointStyle(point?.FillColor, point?.NoFill ?? false,
                point?.Hatch, point?.HatchColor, point?.OutlineColor ?? color,
                point?.OutlineWidth ?? width, point?.ShowOutline ?? visible);
        }
        return result;
    }
}
