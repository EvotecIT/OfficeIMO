using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        internal static bool TryReadPictureEffects(A.Blip blip,
            out LegacyPptWriterPictureEffects effects,
            out string? reason) {
            if (blip == null) throw new ArgumentNullException(nameof(blip));
            A.LuminanceEffect? luminance = null;
            bool grayscale = false;
            bool biLevel = false;
            OfficeColor? transparentColor = null;
            OfficeColor? recolorColor = null;

            foreach (OpenXmlElement child in blip.ChildElements) {
                switch (child) {
                    case A.LuminanceEffect value:
                        if (luminance != null || value.HasChildren
                            || value.ExtendedAttributes.Any()) {
                            effects = LegacyPptWriterPictureEffects.None;
                            reason = "The picture contains a duplicate or extended luminance effect that has no single OfficeArt equivalent.";
                            return false;
                        }
                        int brightness = value.Brightness?.Value ?? 0;
                        int contrast = value.Contrast?.Value ?? 0;
                        if (!IsPictureAdjustment(brightness)
                            || !IsPictureAdjustment(contrast)) {
                            effects = LegacyPptWriterPictureEffects.None;
                            reason = "Picture brightness and contrast must be between -100000 and 100000 for binary PowerPoint.";
                            return false;
                        }
                        luminance = value;
                        break;
                    case A.Grayscale value:
                        if (grayscale || value.HasAttributes
                            || value.HasChildren) {
                            effects = LegacyPptWriterPictureEffects.None;
                            reason = "The picture contains a duplicate or extended grayscale effect that has no single OfficeArt equivalent.";
                            return false;
                        }
                        grayscale = true;
                        break;
                    case A.BiLevel value:
                        if (biLevel || value.HasChildren
                            || value.ExtendedAttributes.Any()
                            || value.Threshold?.Value != 50000) {
                            effects = LegacyPptWriterPictureEffects.None;
                            reason = "Binary PowerPoint supports the classic 50-percent bi-level picture effect only.";
                            return false;
                        }
                        biLevel = true;
                        break;
                    case A.ColorChange value:
                        if (transparentColor.HasValue || value.HasAttributes
                            || value.ColorFrom?.ChildElements.Count != 1
                            || value.ColorTo?.ChildElements.Count != 1
                            || !TryReadEffectRgb(value.ColorFrom?
                                    .GetFirstChild<A.RgbColorModelHex>(),
                                requireTransparent: false,
                                out OfficeColor from)
                            || !TryReadEffectRgb(value.ColorTo?
                                    .GetFirstChild<A.RgbColorModelHex>(),
                                requireTransparent: true,
                                out OfficeColor to)
                            || from != to) {
                            effects = LegacyPptWriterPictureEffects.None;
                            reason = "Binary PowerPoint supports DrawingML color-change only when one RGB source color becomes that same color with zero alpha.";
                            return false;
                        }
                        transparentColor = from;
                        break;
                    case A.ColorReplacement value:
                        if (recolorColor.HasValue || value.HasAttributes
                            || value.ChildElements.Count != 1
                            || !TryReadEffectRgb(value
                                    .GetFirstChild<A.RgbColorModelHex>(),
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
                luminance?.Brightness?.Value ?? 0,
                luminance?.Contrast?.Value ?? 0,
                grayscale, biLevel, transparentColor, recolorColor);
            reason = null;
            return true;
        }

        private static bool IsPictureAdjustment(int value) =>
            value >= -100000 && value <= 100000;

        private static bool TryReadEffectRgb(A.RgbColorModelHex? value,
            bool requireTransparent, out OfficeColor color) {
            color = default;
            string? hex = value?.Val?.Value;
            if (hex == null || value!.ExtendedAttributes.Any()
                || !OfficeColor.TryParse(hex, out color)) {
                return false;
            }
            if (requireTransparent) {
                return value.ChildElements.Count == 1
                    && value.GetFirstChild<A.Alpha>()?.Val?.Value == 0;
            }
            return value.ChildElements.Count == 0;
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
