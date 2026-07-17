using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    private static void AppendTilingPattern(StringBuilder sb, OfficeDrawingTilingPattern pattern, IOfficeRasterImageCodec? imageCodec, ref int gradientId, ref int clipPathId) {
        string clipId = "officeimo-pattern-clip-" + (++clipPathId).ToString(CultureInfo.InvariantCulture);
        OfficeImagePlacement area = pattern.Area;
        sb.Append("<defs><clipPath id=\"").Append(clipId).Append("\"><rect x=\"")
            .Append(Format(area.X)).Append("\" y=\"").Append(Format(area.Y))
            .Append("\" width=\"").Append(Format(area.Width)).Append("\" height=\"")
            .Append(Format(area.Height)).Append("\"/></clipPath></defs><g clip-path=\"url(#")
            .Append(clipId).Append(")\"");
        if (pattern.Opacity < 1D) sb.Append(" opacity=\"").Append(Format(pattern.Opacity)).Append('"');
        sb.Append('>');
        foreach (OfficeTransform transform in pattern.GetTileTransforms(pattern.MaximumTileCount)) {
            sb.Append("<g").Append(BuildMatrixTransformAttribute(transform, 0D, 0D)).Append('>');
            AppendElements(sb, pattern.InnerTile.Elements, imageCodec, ref gradientId, ref clipPathId);
            sb.Append("</g>");
        }
        sb.Append("</g>");
    }
}
