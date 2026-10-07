using System;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgFormatting {
    // Both Drawing and native format adapters supply shrinking-circle fields.
    // Compose RGB and alpha independently so translucent endpoint paint does not
    // accumulate beneath the field. The caller supplies finite visible bounds.
    internal static void AppendNativeRadialPatternDefinition(this StringBuilder builder, string id,
        double left, double top, double width, double height, string outsideColor, string outsideAlpha,
        Action<StringBuilder, string> appendColorField, Action<StringBuilder, string> appendAlphaField) {
        string N(double value) {
            if (double.IsNaN(value) || double.IsInfinity(value)) throw new NotSupportedException("Native radial SVG paint bounds must be finite.");
            return value.ToString("R", CultureInfo.InvariantCulture);
        }
        string region = " x=\"" + N(left) + "\" y=\"" + N(top) + "\" width=\"" + N(width) + "\" height=\"" + N(height) + "\"";
        string colorId = id + "-color", alphaId = id + "-alpha", maskId = id + "-mask";
        builder.Append("<defs><pattern").AppendAttribute("id", id).Append(region)
            .Append(" patternUnits=\"userSpaceOnUse\" patternContentUnits=\"userSpaceOnUse\"><g transform=\"translate(")
            .Append(N(-left)).Append(' ').Append(N(-top)).Append(")\">");
        appendColorField(builder, colorId);
        appendAlphaField(builder, alphaId);
        builder.Append("<defs><mask").AppendAttribute("id", maskId).Append(region)
            .Append(" maskUnits=\"userSpaceOnUse\" maskContentUnits=\"userSpaceOnUse\" style=\"mask-type:luminance\">");
        Rect(outsideAlpha); Rect("url(#" + alphaId + ")");
        builder.Append("</mask></defs><g").AppendAttribute("mask", "url(#" + maskId + ")").Append('>');
        Rect(outsideColor); Rect("url(#" + colorId + ")");
        builder.Append("</g></g></pattern></defs>");
        void Rect(string fill) => builder.Append("<rect").Append(region).AppendAttribute("fill", fill).Append("/>");
    }
}
