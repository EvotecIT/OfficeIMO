using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    private static void AppendEffectGroup(StringBuilder sb, OfficeDrawingEffectGroup effectGroup, IOfficeRasterImageCodec? imageCodec, ref int gradientId, ref int clipPathId) {
        string? maskId = null;
        if (effectGroup.SoftMask != null) {
            maskId = "officeimo-mask-" + (++clipPathId).ToString(CultureInfo.InvariantCulture);
            AppendSoftMaskDefinition(sb, maskId, effectGroup.SoftMask, imageCodec, ref gradientId, ref clipPathId);
        }
        sb.Append("<g").Append(BuildMatrixTransformAttribute(effectGroup.Transform, 0D, 0D));
        if (effectGroup.Opacity < 1D) sb.Append(" opacity=\"").Append(Format(effectGroup.Opacity)).Append('"');
        if (effectGroup.BlendMode != OfficeBlendMode.Normal) sb.Append(" style=\"mix-blend-mode:").Append(ToCssBlendMode(effectGroup.BlendMode)).Append("\"");
        if (maskId != null) sb.Append(" mask=\"url(#").Append(maskId).Append(")\"");
        sb.Append('>');
        AppendElements(sb, effectGroup.InnerDrawing.Elements, imageCodec, ref gradientId, ref clipPathId);
        sb.Append("</g>");
    }

    private static void AppendSoftMaskDefinition(StringBuilder sb, string id, OfficeDrawingSoftMask mask, IOfficeRasterImageCodec? imageCodec, ref int gradientId, ref int clipPathId) {
        sb.Append("<defs><mask id=\"").Append(id)
            .Append("\" maskUnits=\"userSpaceOnUse\" x=\"0\" y=\"0\" width=\"")
            .Append(Format(mask.InnerDrawing.Width)).Append("\" height=\"")
            .Append(Format(mask.InnerDrawing.Height)).Append("\" style=\"mask-type:")
            .Append(mask.Mode == OfficeSoftMaskMode.Alpha ? "alpha" : "luminance")
            .Append("\">");
        if (mask.BackdropColor.A > 0) {
            sb.Append("<rect width=\"100%\" height=\"100%\" fill=\"")
                .Append(mask.BackdropColor.ToHex())
                .Append("\" fill-opacity=\"").Append(Format(mask.BackdropColor.A / 255D)).Append("\"/>");
        }
        sb.Append("<g").Append(BuildMatrixTransformAttribute(mask.Transform, 0D, 0D)).Append('>');
        AppendElements(sb, mask.InnerDrawing.Elements, imageCodec, ref gradientId, ref clipPathId);
        sb.Append("</g></mask></defs>");
    }

    private static string ToCssBlendMode(OfficeBlendMode mode) {
        switch (mode) {
            case OfficeBlendMode.ColorDodge: return "color-dodge";
            case OfficeBlendMode.ColorBurn: return "color-burn";
            case OfficeBlendMode.HardLight: return "hard-light";
            case OfficeBlendMode.SoftLight: return "soft-light";
            default: return mode.ToString().ToLowerInvariant();
        }
    }
}
