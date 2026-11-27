using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;

namespace OfficeIMO.OpenXml.Internal;

internal static partial class OfficeOpenXmlChartSeriesReader {
    private static List<OfficeXmlNode> OrderSeries(List<OfficeXmlNode> elements) {
        // Bounds are established before ordering. Missing order retains legacy XML
        // ordering; explicit orders must be unique to avoid ambiguous projections.
        var orders = new HashSet<uint>();
        foreach (var element in elements) {
            uint? order = element.Child(NsC, "order").UIntVal();
            if (order.HasValue && !orders.Add(order.Value))
                throw new InvalidDataException("Native chart series have duplicate plotting orders.");
        }
        return elements.OrderBy(element => element.Child(NsC, "order").UIntVal() ?? uint.MaxValue).ToList();
    }

    private static bool IsSupportedSeriesShape(OfficeXmlNode properties, OfficeXmlNode scheme, bool filled, bool area = false, bool connectLine = true, bool flattenThreeDimensional = false) {
        if (properties.IsNone) return true;
        foreach (var child in properties.Children()) {
            bool drawing = child.NamespaceId == NsA;
            string name = child.LocalName;
            if (drawing && name is "effectLst" or "sp3d" && !child.HasChildElements() && !child.HasAttributes) continue;
            if (flattenThreeDimensional && drawing && name == "sp3d") continue;
            if (drawing && name == "solidFill" && !OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(child) &&
                OfficeOpenXmlThemeColorResolver.ResolveColor(child, scheme).HasValue) continue;
            if (drawing && name == "noFill") continue;
            if (drawing && name == "pattFill" && filled && OfficeOpenXmlChartPointStyles.IsSupportedPattern(child, scheme)) continue;
            if (!(drawing && name == "ln")) return false;
            OfficeXmlNode outline = child;
            // Marker-only series retain native line metadata that has no rendered
            // effect. Filled series still use the outline around their geometry.
            if (!filled && !connectLine) continue;
            int? lineWidth = outline.IntAttr("w");
            if (lineWidth == 0 && outline.Child(NsA, "noFill").IsNone) return false;
            if (OfficeOpenXmlChartPointStyles.HasUnsupportedLineAttributes(outline) || lineWidth < 0) return false;
            foreach (var lineChild in outline.Children()) {
                bool lineDrawing = lineChild.NamespaceId == NsA;
                string lineName = lineChild.LocalName;
                if (lineDrawing && lineName is "noFill" or "round" or "bevel") continue;
                if (lineDrawing && lineName == "miter" && !lineChild.HasAttributes && !lineChild.HasChildElements()) continue;
                if (lineDrawing && lineName == "solidFill" && !OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(lineChild) &&
                    OfficeOpenXmlThemeColorResolver.ResolveColor(lineChild, scheme).HasValue) continue;
                if ((!filled || area) && lineDrawing && lineName == "prstDash" && (!connectLine || ReadDash(outline).HasValue)) continue;
                return false;
            }
        }
        return true;
    }

}
