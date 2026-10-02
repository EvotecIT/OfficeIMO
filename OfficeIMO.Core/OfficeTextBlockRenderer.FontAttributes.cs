using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextBlockRenderer {
    private static void AppendSvgFontFaceAttributes(StringBuilder builder, OfficeFontFaceDescriptor? face, bool bold, bool italic) {
        if (face.HasValue) {
            builder.Append(" font-weight=\"").Append(face.Value.Weight.ToString(CultureInfo.InvariantCulture)).Append('"');
            if (face.Value.StretchPercent != 100D) builder.AppendAttribute("font-stretch", face.Value.StretchPercent.ToString(CultureInfo.InvariantCulture) + "%");
            if (face.Value.Slant == OfficeFontSlant.Oblique) builder.AppendAttribute("font-style", "oblique " + face.Value.ObliqueAngleDegrees.ToString(CultureInfo.InvariantCulture) + "deg");
            else if (face.Value.Slant == OfficeFontSlant.Italic) builder.Append(" font-style=\"italic\"");
        } else {
            if (bold) builder.Append(" font-weight=\"700\"");
            if (italic) builder.Append(" font-style=\"italic\"");
        }
    }
}
