using System;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool TryResolveSupportedRasterTransform(
        XElement element,
        OfficeTransform inherited,
        double viewX,
        double viewY,
        out OfficeTransform transform) {
        string? value = ReadRasterProjectedAttribute(element, "transform");
        if (string.IsNullOrWhiteSpace(value)) {
            transform = inherited;
            return true;
        }
        if (!OfficeSvgTransformParser.TryParse(value, out OfficeTransform parsed)) {
            transform = inherited;
            return false;
        }

        OfficeTransform normalized = OfficeTransform.Translate(viewX, viewY)
            .Then(parsed)
            .Then(OfficeTransform.Translate(-viewX, -viewY));
        transform = normalized.Then(inherited);
        return IsSupportedSvgTransform(transform);
    }
}
