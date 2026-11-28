using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using static OfficeIMO.PowerPoint.LegacyPpt.Internal.LegacyPptXml;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const uint TextAutoMarginMasks = (1U << 12) | (1U << 28);
        private const uint TextFitShapeMasks = (1U << 14) | (1U << 30);

        /// <summary>Reads the text frame of a text shape from the presentation captured for the write.</summary>
        internal static bool TryReadTextFrameForWrite(
            LegacyPptPackageView view, PowerPointTextBox textBox,
            out IReadOnlyList<LegacyPptWriterFoptProperty> properties,
            out string? reason) {
            ArgumentNullException.ThrowIfNull(view);
            return TryReadTextFrameForWrite(ResolveShapeForWrite(view, textBox),
                out properties, out reason);
        }

        /// <summary>Reads the text frame of a text shape as the document holds it now.</summary>
        internal static bool TryReadTextFrameForWrite(
            PowerPointTextBox textBox,
            out IReadOnlyList<LegacyPptWriterFoptProperty> properties,
            out string? reason) => TryReadTextFrameForWrite(
                textBox == null ? default : textBox.XmlElement,
                out properties, out reason);

        /// <summary>Reads the text frame (<c>p:txBody/a:bodyPr</c>) of a <c>p:sp</c> element.</summary>
        internal static bool TryReadTextFrameForWrite(
            OfficeXmlNode shape,
            out IReadOnlyList<LegacyPptWriterFoptProperty> properties,
            out string? reason) {
            properties = Array.Empty<LegacyPptWriterFoptProperty>();
            reason = null;
            OfficeXmlNode body = shape.Is(P, "sp")
                ? shape.Path(P, "txBody").Child(A, "bodyPr")
                : default;
            if (body.IsNone) {
                reason = "The text shape has no DrawingML body properties.";
                return false;
            }
            foreach (string name in AttributeLocalNames(body)) {
                if (name is not "vert" and not "wrap" and not "lIns"
                    and not "tIns" and not "rIns" and not "bIns"
                    and not "anchor" and not "anchorCtr") {
                    reason = "The text frame contains a DrawingML attribute with no exact classic OfficeArt mapping.";
                    return false;
                }
            }
            int fitCount = 0;
            foreach (OfficeXmlNode child in body.Children()) {
                if (!child.Is(A, "noAutofit") && !child.Is(A, "spAutoFit")
                    && !child.Is(A, "normAutofit")) {
                    fitCount = int.MaxValue;
                    break;
                }
                fitCount++;
            }
            if (fitCount > 1) {
                reason = "The text frame contains unsupported or duplicate body-property children.";
                return false;
            }
            if (!body.Child(A, "normAutofit").IsNone) {
                reason = "Shrink-text normal autofit has no lossless classic binary PowerPoint mapping.";
                return false;
            }
            foreach (OfficeXmlNode fit in body.Children()) {
                if (HasAnyAttribute(fit) || HasChildElements(fit)) {
                    reason = "The text autofit element contains an unsupported extension.";
                    return false;
                }
            }

            var result = new List<LegacyPptWriterFoptProperty>(9);
            if (!TryAddTextInset(result, 0x0081, body.Int("lIns"),
                    out reason)
                || !TryAddTextInset(result, 0x0082,
                    body.Int("tIns"), out reason)
                || !TryAddTextInset(result, 0x0083,
                    body.Int("rIns"), out reason)
                || !TryAddTextInset(result, 0x0084,
                    body.Int("bIns"), out reason)) {
                return false;
            }
            // ST_TextWrappingType has only these two tokens; anything else reads as absent.
            string? wrap = body.Attribute("wrap");
            if (wrap is "square" or "none") {
                result.Add(new LegacyPptWriterFoptProperty(0x0085,
                    wrap == "square" ? 0U : 2U));
            }
            string? anchorToken = body.Attribute("anchor");
            bool hasAnchor = anchorToken is "t" or "ctr" or "b" or "just" or "dist";
            bool? anchorCenter = body.Bool("anchorCtr");
            if (hasAnchor || anchorCenter.HasValue) {
                uint anchor = hasAnchor && anchorToken == "ctr" ? 1U
                    : hasAnchor && anchorToken == "b" ? 2U : 0U;
                if (anchorCenter == true) anchor += 3U;
                result.Add(new LegacyPptWriterFoptProperty(0x0087,
                    anchor));
            }
            string? vertical = body.Attribute("vert");
            if (vertical is "horz" or "vert" or "vert270" or "wordArtVert"
                or "eaVert" or "mongolianVert" or "wordArtVertRtl") {
                uint flow = vertical == "horz" ? 0U
                    : vertical == "vert" ? 1U
                    : vertical == "vert270" ? 2U
                    : uint.MaxValue;
                if (flow == uint.MaxValue) {
                    reason = "The text direction has no exact classic OfficeArt mapping.";
                    return false;
                }
                result.Add(new LegacyPptWriterFoptProperty(0x0088, flow));
            }

            uint textFlags = 0U;
            if (HasExplicitTextInsets(body)) textFlags |= 1U << 12;
            if (!body.Child(A, "noAutofit").IsNone) {
                textFlags |= 1U << 14;
            } else if (!body.Child(A, "spAutoFit").IsNone) {
                textFlags |= (1U << 14) | (1U << 30);
            }
            if (textFlags != 0U) {
                result.Add(new LegacyPptWriterFoptProperty(0x00BF,
                    textFlags));
            }
            properties = result;
            return true;
        }

        /// <summary>Whether <c>a:bodyPr</c> sets any of its four insets.</summary>
        internal static bool HasExplicitTextInsets(OfficeXmlNode body) =>
            body.Int("lIns").HasValue
            || body.Int("tIns").HasValue
            || body.Int("rIns").HasValue
            || body.Int("bIns").HasValue;

        /// <summary>
        /// A shape's element in <paramref name="view"/>, or as the document holds it now when there is no view. Every
        /// element the writer compares by identity (picture bullets) comes from one of the two, never both.
        /// </summary>
        private static OfficeXmlNode ResolveShapeForWrite(
            LegacyPptPackageView? view, PowerPointShape? shape) {
            if (shape == null) return default;
            OfficeXmlNode live = shape.XmlElement;
            return view == null ? live : view.Resolve(shape.OwnerPartRef.Name, live);
        }

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
