using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const uint TextAutoMarginMasks = (1U << 12) | (1U << 28);
        private const uint TextFitShapeMasks = (1U << 14) | (1U << 30);

        internal static bool TryReadTextFrameForWrite(
            PowerPointTextBox textBox,
            out IReadOnlyList<LegacyPptWriterFoptProperty> properties,
            out string? reason) {
            properties = Array.Empty<LegacyPptWriterFoptProperty>();
            reason = null;
            if (textBox?.Element is not P.Shape shape
                || shape.TextBody?.BodyProperties is not A.BodyProperties
                    body) {
                reason = "The text shape has no DrawingML body properties.";
                return false;
            }
            if (body.GetAttributes().Any(attribute => attribute.LocalName
                    is not "vert" and not "wrap" and not "lIns"
                    and not "tIns" and not "rIns" and not "bIns"
                    and not "anchor" and not "anchorCtr")) {
                reason = "The text frame contains a DrawingML attribute with no exact classic OfficeArt mapping.";
                return false;
            }
            if (body.ChildElements.Any(child => child is not A.NoAutoFit
                    and not A.ShapeAutoFit and not A.NormalAutoFit)
                || body.Elements<A.NoAutoFit>().Count()
                    + body.Elements<A.ShapeAutoFit>().Count()
                    + body.Elements<A.NormalAutoFit>().Count() > 1) {
                reason = "The text frame contains unsupported or duplicate body-property children.";
                return false;
            }
            A.NormalAutoFit? normal = body.GetFirstChild<A.NormalAutoFit>();
            if (normal != null) {
                reason = "Shrink-text normal autofit has no lossless classic binary PowerPoint mapping.";
                return false;
            }
            foreach (OpenXmlElement fit in body.ChildElements) {
                if (fit.HasAttributes || fit.HasChildren) {
                    reason = "The text autofit element contains an unsupported extension.";
                    return false;
                }
            }

            var result = new List<LegacyPptWriterFoptProperty>(9);
            if (!TryAddTextInset(result, 0x0081, body.LeftInset?.Value,
                    out reason)
                || !TryAddTextInset(result, 0x0082,
                    body.TopInset?.Value, out reason)
                || !TryAddTextInset(result, 0x0083,
                    body.RightInset?.Value, out reason)
                || !TryAddTextInset(result, 0x0084,
                    body.BottomInset?.Value, out reason)) {
                return false;
            }
            if (body.Wrap?.HasValue == true) {
                A.TextWrappingValues value = body.Wrap.Value;
                uint wrap = value == A.TextWrappingValues.Square ? 0U
                    : value == A.TextWrappingValues.None ? 2U
                    : uint.MaxValue;
                if (wrap == uint.MaxValue) {
                    reason = "The text wrapping mode has no exact classic OfficeArt mapping.";
                    return false;
                }
                result.Add(new LegacyPptWriterFoptProperty(0x0085, wrap));
            }
            if (body.Anchor?.HasValue == true
                || body.AnchorCenter?.HasValue == true) {
                A.TextAnchoringTypeValues? value = body.Anchor?.Value;
                uint anchor = value == A.TextAnchoringTypeValues.Center ? 1U
                    : value == A.TextAnchoringTypeValues.Bottom ? 2U : 0U;
                if (body.AnchorCenter?.Value == true) anchor += 3U;
                result.Add(new LegacyPptWriterFoptProperty(0x0087,
                    anchor));
            }
            if (body.Vertical?.HasValue == true) {
                A.TextVerticalValues value = body.Vertical.Value;
                uint flow = value == A.TextVerticalValues.Horizontal ? 0U
                    : value == A.TextVerticalValues.Vertical ? 1U
                    : value == A.TextVerticalValues.Vertical270 ? 2U
                    : uint.MaxValue;
                if (flow == uint.MaxValue) {
                    reason = "The text direction has no exact classic OfficeArt mapping.";
                    return false;
                }
                result.Add(new LegacyPptWriterFoptProperty(0x0088, flow));
            }

            uint textFlags = 0U;
            if (HasExplicitTextInsets(body)) textFlags |= 1U << 12;
            if (body.GetFirstChild<A.NoAutoFit>() != null) {
                textFlags |= 1U << 14;
            } else if (body.GetFirstChild<A.ShapeAutoFit>() != null) {
                textFlags |= (1U << 14) | (1U << 30);
            }
            if (textFlags != 0U) {
                result.Add(new LegacyPptWriterFoptProperty(0x00BF,
                    textFlags));
            }
            properties = result;
            return true;
        }

        internal static bool HasExplicitTextInsets(A.BodyProperties body) =>
            body.LeftInset?.HasValue == true
            || body.TopInset?.HasValue == true
            || body.RightInset?.HasValue == true
            || body.BottomInset?.HasValue == true;

        private static bool TryAddTextInset(
            ICollection<LegacyPptWriterFoptProperty> properties,
            ushort propertyId, int? value, out string? reason) {
            reason = null;
            if (!value.HasValue) return true;
            if (value.Value < 0 || value.Value > 0x0132F540) {
                reason = "A text inset lies outside the classic OfficeArt coordinate range.";
                return false;
            }
            properties.Add(new LegacyPptWriterFoptProperty(propertyId,
                checked((uint)value.Value)));
            return true;
        }
    }
}
