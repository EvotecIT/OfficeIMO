using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using static OfficeIMO.PowerPoint.LegacyPpt.Write.LegacyPptWriterXml;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        internal static bool TryReadPictureEffects(OfficeXmlNode blip,
            out LegacyPptWriterPictureEffects effects,
            out string? reason) {
            if (blip.IsNone) throw new ArgumentNullException(nameof(blip));
            OfficeXmlNode luminance = default;
            bool grayscale = false;
            bool biLevel = false;
            OfficeColor? transparentColor = null;
            OfficeColor? recolorColor = null;

            foreach (OfficeXmlNode child in blip.Children()) {
                string name = child.NamespaceId == LegacyPptXml.A
                    ? child.LocalName : string.Empty;
                switch (name) {
                    case "lum":
                        if (!luminance.IsNone || HasChildren(child)
                            || HasExtendedAttributes(child, "bright",
                                "contrast")) {
                            effects = LegacyPptWriterPictureEffects.None;
                            reason = "The picture contains a duplicate or extended luminance effect that has no single OfficeArt equivalent.";
                            return false;
                        }
                        int brightness = child.Int("bright") ?? 0;
                        int contrast = child.Int("contrast") ?? 0;
                        if (!IsPictureAdjustment(brightness)
                            || !IsPictureAdjustment(contrast)) {
                            effects = LegacyPptWriterPictureEffects.None;
                            reason = "Picture brightness and contrast must be between -100000 and 100000 for binary PowerPoint.";
                            return false;
                        }
                        luminance = child;
                        break;
                    case "grayscl":
                        if (grayscale || HasAttributes(child)
                            || HasChildren(child)) {
                            effects = LegacyPptWriterPictureEffects.None;
                            reason = "The picture contains a duplicate or extended grayscale effect that has no single OfficeArt equivalent.";
                            return false;
                        }
                        grayscale = true;
                        break;
                    case "biLevel":
                        if (biLevel || HasChildren(child)
                            || HasExtendedAttributes(child, "thresh")
                            || child.Int("thresh") != 50000) {
                            effects = LegacyPptWriterPictureEffects.None;
                            reason = "Binary PowerPoint supports the classic 50-percent bi-level picture effect only.";
                            return false;
                        }
                        biLevel = true;
                        break;
                    case "clrChange":
                        OfficeXmlNode colorFrom = child.Child(LegacyPptXml.A,
                            "clrFrom");
                        OfficeXmlNode colorTo = child.Child(LegacyPptXml.A,
                            "clrTo");
                        if (transparentColor.HasValue || HasAttributes(child)
                            || colorFrom.IsNone || ChildCount(colorFrom) != 1
                            || colorTo.IsNone || ChildCount(colorTo) != 1
                            || !TryReadEffectRgb(colorFrom.Child(
                                    LegacyPptXml.A, "srgbClr"),
                                requireTransparent: false,
                                out OfficeColor from)
                            || !TryReadEffectRgb(colorTo.Child(
                                    LegacyPptXml.A, "srgbClr"),
                                requireTransparent: true,
                                out OfficeColor to)
                            || from != to) {
                            effects = LegacyPptWriterPictureEffects.None;
                            reason = "Binary PowerPoint supports DrawingML color-change only when one RGB source color becomes that same color with zero alpha.";
                            return false;
                        }
                        transparentColor = from;
                        break;
                    case "clrRepl":
                        if (recolorColor.HasValue || HasAttributes(child)
                            || ChildCount(child) != 1
                            || !TryReadEffectRgb(child.Child(LegacyPptXml.A,
                                    "srgbClr"),
                                requireTransparent: false,
                                out OfficeColor replacement)) {
                            effects = LegacyPptWriterPictureEffects.None;
                            reason = "Binary PowerPoint picture recoloring requires one unmodified RGB replacement color.";
                            return false;
                        }
                        recolorColor = replacement;
                        break;
                    default:
                        effects = LegacyPptWriterPictureEffects.None;
                        reason = $"The DrawingML picture effect '{child.LocalName}' has no native OfficeArt picture-property mapping.";
                        return false;
                }
            }
            if (grayscale && biLevel) {
                effects = LegacyPptWriterPictureEffects.None;
                reason = "A DrawingML grayscale and bi-level effect sequence cannot be reduced to unordered classic picture flags without changing its result.";
                return false;
            }

            effects = new LegacyPptWriterPictureEffects(
                luminance.IsNone ? 0 : luminance.Int("bright") ?? 0,
                luminance.IsNone ? 0 : luminance.Int("contrast") ?? 0,
                grayscale, biLevel, transparentColor, recolorColor);
            reason = null;
            return true;
        }

        private static bool IsPictureAdjustment(int value) =>
            value >= -100000 && value <= 100000;

        private static bool TryReadEffectRgb(OfficeXmlNode value,
            bool requireTransparent, out OfficeColor color) {
            color = default;
            string? hex = value.IsNone ? null : value.Attribute("val");
            if (hex == null || HasExtendedAttributes(value, "val")
                || !OfficeColor.TryParse(hex, out color)) {
                return false;
            }
            if (requireTransparent) {
                return ChildCount(value) == 1
                    && value.Child(LegacyPptXml.A, "alpha").Int("val") == 0;
            }
            return ChildCount(value) == 0;
        }

        private static void AddPictureEffectProperties(
            ICollection<LegacyPptWriterFoptProperty> properties,
            LegacyPptWriterPictureEffects effects,
            uint preservedBooleanProperties = 0) {
            if (effects.Contrast != 0) {
                properties.Add(new LegacyPptWriterFoptProperty(0x0108,
                    unchecked((uint)ToOfficeArtContrast(effects.Contrast))));
            }
            if (effects.Brightness != 0) {
                properties.Add(new LegacyPptWriterFoptProperty(0x0109,
                    unchecked((uint)ToOfficeArtBrightness(
                        effects.Brightness))));
            }
            if (effects.TransparentColor.HasValue) {
                properties.Add(new LegacyPptWriterFoptProperty(0x0107,
                    PackOfficeArtColor(effects.TransparentColor.Value)));
            }
            uint booleanProperties = preservedBooleanProperties;
            // PowerPoint sets the grayscale bit together with bi-level. The
            // latter selects the black/white rendering mode while the former
            // activates the classic color-conversion path.
            if (effects.Grayscale || effects.BiLevel) {
                booleanProperties |= 1U << 18;
                booleanProperties |= 1U << 2;
            }
            if (effects.BiLevel) {
                booleanProperties |= 1U << 17;
                booleanProperties |= 1U << 1;
            }
            if (booleanProperties != 0) {
                properties.Add(new LegacyPptWriterFoptProperty(0x013F,
                    booleanProperties));
            }
        }

        private static int ToOfficeArtBrightness(int value) => checked((int)
            Math.Round(value / 100000D * 32768D,
                MidpointRounding.AwayFromZero));

        private static int ToOfficeArtContrast(int value) {
            if (value >= 100000) return int.MaxValue;
            double adjustment = value / 100000D;
            double raw = adjustment <= 0D
                ? (1D + adjustment) * 65536D
                : 65536D / (1D - adjustment);
            return checked((int)Math.Truncate(raw));
        }

        internal sealed class LegacyPptWriterPictureEffects {
            internal static LegacyPptWriterPictureEffects None { get; } =
                new(0, 0, grayscale: false, biLevel: false,
                    transparentColor: null, recolorColor: null);

            internal LegacyPptWriterPictureEffects(int brightness,
                int contrast, bool grayscale, bool biLevel,
                OfficeColor? transparentColor, OfficeColor? recolorColor) {
                Brightness = brightness;
                Contrast = contrast;
                Grayscale = grayscale;
                BiLevel = biLevel;
                TransparentColor = transparentColor;
                RecolorColor = recolorColor;
            }

            internal int Brightness { get; }
            internal int Contrast { get; }
            internal bool Grayscale { get; }
            internal bool BiLevel { get; }
            internal OfficeColor? TransparentColor { get; }
            internal OfficeColor? RecolorColor { get; }
        }
    }
}
