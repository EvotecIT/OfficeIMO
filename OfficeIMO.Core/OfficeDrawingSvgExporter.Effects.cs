using System;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    private static void AppendEffectGroup(StringBuilder sb, OfficeDrawingEffectGroup effectGroup, IOfficeRasterImageCodec? imageCodec, string idPrefix, ref int gradientId, ref int clipPathId, System.Threading.CancellationToken cancellationToken, SvgTilingExpansionBudget tilingExpansionBudget, OfficeRasterCanvas textMetrics, bool textAsPaths, System.Collections.Generic.ISet<string> requiredFontFamilies) {
        if (effectGroup.Opacity <= 0D) return;
        textMetrics = textMetrics.WithDrawingTextProfile(effectGroup.InnerDrawing);
        string? maskId = null;
        if (effectGroup.SoftMask != null) {
            maskId = idPrefix + "officeimo-mask-" + (++clipPathId).ToString(CultureInfo.InvariantCulture);
            AppendSoftMaskDefinition(sb, maskId, effectGroup.SoftMask, imageCodec, idPrefix, ref gradientId, ref clipPathId, cancellationToken, tilingExpansionBudget, textMetrics, textAsPaths, requiredFontFamilies);
        }
        sb.Append("<g").Append(BuildMatrixTransformAttribute(effectGroup.Transform, 0D, 0D));
        if (effectGroup.Opacity < 1D) sb.Append(" opacity=\"").Append(Format(effectGroup.Opacity)).Append('"');
        sb.Append(" style=\"isolation:isolate");
        if (effectGroup.BlendMode != OfficeBlendMode.Normal) sb.Append(";mix-blend-mode:").Append(ToCssBlendMode(effectGroup.BlendMode));
        sb.Append('"');
        if (maskId != null) sb.Append(" mask=\"url(#").Append(maskId).Append(")\"");
        sb.Append('>');
        AppendElements(sb, effectGroup.InnerDrawing.Elements, imageCodec, idPrefix, ref gradientId, ref clipPathId, cancellationToken, tilingExpansionBudget, textMetrics, textAsPaths, requiredFontFamilies);
        sb.Append("</g>");
    }

    private static void AppendSoftMaskDefinition(StringBuilder sb, string id, OfficeDrawingSoftMask mask, IOfficeRasterImageCodec? imageCodec, string idPrefix, ref int gradientId, ref int clipPathId, System.Threading.CancellationToken cancellationToken, SvgTilingExpansionBudget tilingExpansionBudget, OfficeRasterCanvas textMetrics, bool textAsPaths, System.Collections.Generic.ISet<string> requiredFontFamilies) {
        textMetrics = textMetrics.WithDrawingTextProfile(mask.InnerDrawing);
        bool hasTransfer = mask.TransferTable != null;
        bool useFilter = mask.Mode == OfficeSoftMaskMode.Alpha || hasTransfer ||
            mask.LuminosityStandard == OfficeSoftMaskLuminosityStandard.PdfDeviceRgb;
        string filterId = id + "-coverage";
        sb.Append("<defs>");
        if (useFilter) AppendSoftMaskCoverageFilter(sb, filterId, mask, cancellationToken);
        sb.Append("<mask id=\"").Append(id)
            .Append("\" maskUnits=\"userSpaceOnUse\" x=\"0\" y=\"0\" width=\"")
            .Append(Format(mask.InnerDrawing.Width)).Append("\" height=\"")
            .Append(Format(mask.InnerDrawing.Height)).Append("\" style=\"mask-type:")
            .Append("luminance")
            .Append("\">");
        if (useFilter) sb.Append("<g filter=\"url(#").Append(filterId).Append(")\">");
        if (hasTransfer) {
            sb.Append("<rect").AppendNumberAttribute("width", mask.InnerDrawing.Width)
                .AppendNumberAttribute("height", mask.InnerDrawing.Height)
                .AppendAttribute("fill", "transparent").Append("/>");
        }
        if (mask.BackdropColor.A > 0) {
            sb.Append("<rect width=\"100%\" height=\"100%\" fill=\"#")
                .Append(mask.BackdropColor.ToRgbHex())
                .Append("\" fill-opacity=\"").Append(Format(mask.BackdropColor.A / 255D)).Append("\"/>");
        }
        sb.Append("<g").Append(BuildMatrixTransformAttribute(mask.Transform, 0D, 0D)).Append('>');
        AppendElements(sb, mask.InnerDrawing.Elements, imageCodec, idPrefix, ref gradientId, ref clipPathId, cancellationToken, tilingExpansionBudget, textMetrics, textAsPaths, requiredFontFamilies);
        sb.Append("</g>");
        if (useFilter) sb.Append("</g>");
        sb.Append("</mask></defs>");
    }

    // Apply /TR to the final mask coverage, after luminosity and source alpha.
    // An explicit filter region also evaluates T(0) in otherwise empty mask areas.
    private static void AppendSoftMaskCoverageFilter(StringBuilder sb, string id, OfficeDrawingSoftMask mask,
        System.Threading.CancellationToken cancellationToken) {
        byte[]? transfer = mask.TransferTable;
        if (transfer != null && transfer.Length != 256) {
            throw new InvalidOperationException("A soft-mask transfer table must contain 256 samples.");
        }
        sb.Append("<filter").AppendAttribute("id", id)
            .AppendAttribute("filterUnits", "userSpaceOnUse")
            .AppendNumberAttribute("x", 0D).AppendNumberAttribute("y", 0D)
            .AppendNumberAttribute("width", mask.InnerDrawing.Width)
            .AppendNumberAttribute("height", mask.InnerDrawing.Height)
            .AppendAttribute("color-interpolation-filters", "sRGB").Append('>');
        if (mask.Mode == OfficeSoftMaskMode.Luminosity) {
            string weights = mask.LuminosityStandard == OfficeSoftMaskLuminosityStandard.PdfDeviceRgb
                ? "0.3 0.59 0.11" : "0.2126 0.7152 0.0722";
            sb.Append("<feColorMatrix type=\"matrix\" values=\"0 0 0 0 1 0 0 0 0 1 0 0 0 0 1 ")
                .Append(weights).Append(" 0 0\"/>")
                .Append("<feComposite in2=\"SourceAlpha\" operator=\"in\"/>");
        }
        // Put coverage in RGB with opaque alpha. A luminance mask then has the
        // same meaning in SVG 1.1 renderers and engines supporting mask-type.
        // Alpha-only transfer cannot reveal transparent black in a luminance mask.
        sb.Append("<feColorMatrix type=\"matrix\" values=\"0 0 0 1 0 0 0 0 1 0 0 0 0 1 0 0 0 0 0 1\"/>");
        if (transfer != null) {
            sb.Append("<feComponentTransfer>");
            foreach (string channel in new[] { "R", "G", "B" }) {
                // One discrete entry per byte-domain coverage sample. Keeping at
                // most 256 entries also supports SVG renderers with bounded tables.
                sb.Append("<feFunc").Append(channel).Append(" type=\"discrete\" tableValues=\"");
                for (int index = 0; index < 256; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (index > 0) sb.Append(' ');
                    sb.Append((transfer[index] / 255D).ToString("0.########", CultureInfo.InvariantCulture));
                }
                sb.Append("\"/>");
            }
            sb.Append("</feComponentTransfer>");
        }
        sb.Append("</filter>");
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
