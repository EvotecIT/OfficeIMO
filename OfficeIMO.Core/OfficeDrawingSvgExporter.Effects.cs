using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    private static void AppendEffectGroup(StringBuilder sb, OfficeDrawingEffectGroup effectGroup, IOfficeRasterImageCodec? imageCodec, string idPrefix, ref int gradientId, ref int clipPathId, System.Threading.CancellationToken cancellationToken, SvgTilingExpansionBudget tilingExpansionBudget) {
        if (effectGroup.Opacity <= 0D) return;
        string? maskId = null;
        if (effectGroup.SoftMask != null) {
            maskId = idPrefix + "officeimo-mask-" + (++clipPathId).ToString(CultureInfo.InvariantCulture);
            AppendSoftMaskDefinition(sb, maskId, effectGroup.SoftMask, imageCodec, idPrefix, ref gradientId, ref clipPathId, cancellationToken, tilingExpansionBudget);
        }
        sb.Append("<g").Append(BuildMatrixTransformAttribute(effectGroup.Transform, 0D, 0D));
        if (effectGroup.Opacity < 1D) sb.Append(" opacity=\"").Append(Format(effectGroup.Opacity)).Append('"');
        sb.Append(" style=\"isolation:isolate");
        if (effectGroup.BlendMode != OfficeBlendMode.Normal) sb.Append(";mix-blend-mode:").Append(ToCssBlendMode(effectGroup.BlendMode));
        sb.Append('"');
        if (maskId != null) sb.Append(" mask=\"url(#").Append(maskId).Append(")\"");
        sb.Append('>');
        AppendElements(sb, effectGroup.InnerDrawing.Elements, imageCodec, idPrefix, ref gradientId, ref clipPathId, cancellationToken, tilingExpansionBudget);
        sb.Append("</g>");
    }

    private static void AppendSoftMaskDefinition(StringBuilder sb, string id, OfficeDrawingSoftMask mask, IOfficeRasterImageCodec? imageCodec, string idPrefix, ref int gradientId, ref int clipPathId, System.Threading.CancellationToken cancellationToken, SvgTilingExpansionBudget tilingExpansionBudget) {
        bool pdfLuminosity = mask.Mode == OfficeSoftMaskMode.Luminosity &&
            mask.LuminosityStandard == OfficeSoftMaskLuminosityStandard.PdfDeviceRgb;
        string filterId = id + "-pdf-luminosity";
        sb.Append("<defs>");
        if (pdfLuminosity) {
            sb.Append("<filter id=\"").Append(filterId)
                .Append("\" color-interpolation-filters=\"sRGB\"><feColorMatrix type=\"matrix\" values=\"0.3 0.59 0.11 0 0 0.3 0.59 0.11 0 0 0.3 0.59 0.11 0 0 0 0 0 1 0\"/></filter>");
        }
        sb.Append("<mask id=\"").Append(id)
            .Append("\" maskUnits=\"userSpaceOnUse\" x=\"0\" y=\"0\" width=\"")
            .Append(Format(mask.InnerDrawing.Width)).Append("\" height=\"")
            .Append(Format(mask.InnerDrawing.Height)).Append("\" style=\"mask-type:")
            .Append(mask.Mode == OfficeSoftMaskMode.Alpha ? "alpha" : "luminance")
            .Append("\">");
        if (pdfLuminosity) sb.Append("<g filter=\"url(#").Append(filterId).Append(")\">");
        if (mask.BackdropColor.A > 0) {
            sb.Append("<rect width=\"100%\" height=\"100%\" fill=\"")
                .Append(mask.BackdropColor.ToHex())
                .Append("\" fill-opacity=\"").Append(Format(mask.BackdropColor.A / 255D)).Append("\"/>");
        }
        sb.Append("<g").Append(BuildMatrixTransformAttribute(mask.Transform, 0D, 0D)).Append('>');
        AppendElements(sb, mask.InnerDrawing.Elements, imageCodec, idPrefix, ref gradientId, ref clipPathId, cancellationToken, tilingExpansionBudget);
        sb.Append("</g>");
        if (pdfLuminosity) sb.Append("</g>");
        sb.Append("</mask></defs>");
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
