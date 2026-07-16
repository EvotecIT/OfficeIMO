using A = DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointPicture {
        /// <summary>
        /// Gets or sets the luminance brightness adjustment as a percentage
        /// from -100 through 100, or <see langword="null"/> when omitted.
        /// </summary>
        public int? LuminanceBrightness {
            get => Picture.BlipFill?.Blip?
                .GetFirstChild<A.LuminanceEffect>()?.Brightness?.Value / 1000;
            set {
                ValidateEffectPercent(value, nameof(LuminanceBrightness));
                UpdateLuminance(value, LuminanceContrast);
            }
        }

        /// <summary>
        /// Gets or sets the luminance contrast adjustment as a percentage
        /// from -100 through 100, or <see langword="null"/> when omitted.
        /// </summary>
        public int? LuminanceContrast {
            get => Picture.BlipFill?.Blip?
                .GetFirstChild<A.LuminanceEffect>()?.Contrast?.Value / 1000;
            set {
                ValidateEffectPercent(value, nameof(LuminanceContrast));
                UpdateLuminance(LuminanceBrightness, value);
            }
        }

        /// <summary>Gets or sets whether the picture uses grayscale display.</summary>
        public bool GrayScale {
            get => Picture.BlipFill?.Blip?.GetFirstChild<A.Grayscale>()
                != null;
            set {
                A.Blip blip = GetRequiredBlip();
                A.Grayscale? grayscale = blip.GetFirstChild<A.Grayscale>();
                if (value && grayscale == null) {
                    blip.Append(new A.Grayscale());
                } else if (!value) {
                    grayscale?.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets the black-and-white threshold as a percentage from
        /// 0 through 100, or <see langword="null"/> to disable bi-level
        /// display. Binary PowerPoint supports the classic value 50.
        /// </summary>
        public int? BlackWhiteThreshold {
            get => Picture.BlipFill?.Blip?.GetFirstChild<A.BiLevel>()?
                .Threshold?.Value / 1000;
            set {
                if (value is < 0 or > 100) {
                    throw new ArgumentOutOfRangeException(nameof(value),
                        "Threshold must be between 0 and 100.");
                }
                A.Blip blip = GetRequiredBlip();
                A.BiLevel? biLevel = blip.GetFirstChild<A.BiLevel>();
                if (!value.HasValue) {
                    biLevel?.Remove();
                    return;
                }
                if (biLevel == null) {
                    biLevel = new A.BiLevel();
                    blip.Append(biLevel);
                }
                biLevel.Threshold = value.Value * 1000;
            }
        }

        /// <summary>
        /// Gets or sets the RGB color that becomes fully transparent, or
        /// <see langword="null"/> to disable transparent-color replacement.
        /// </summary>
        public OfficeColor? TransparentColor {
            get {
                A.ColorChange? change = Picture.BlipFill?.Blip?
                    .GetFirstChild<A.ColorChange>();
                A.RgbColorModelHex? from = change?.ColorFrom?
                    .GetFirstChild<A.RgbColorModelHex>();
                A.RgbColorModelHex? to = change?.ColorTo?
                    .GetFirstChild<A.RgbColorModelHex>();
                if (from?.Val?.Value == null || to?.Val?.Value == null
                    || from.ChildElements.Count != 0
                    || to.ChildElements.Count != 1
                    || !string.Equals(from.Val.Value, to.Val.Value,
                        StringComparison.OrdinalIgnoreCase)
                    || to.GetFirstChild<A.Alpha>()?.Val?.Value != 0
                    || !OfficeColor.TryParse(from.Val.Value,
                        out OfficeColor color)) {
                    return null;
                }
                return color;
            }
            set {
                A.Blip blip = GetRequiredBlip();
                A.ColorChange? change = blip.GetFirstChild<A.ColorChange>();
                if (!value.HasValue) {
                    change?.Remove();
                    return;
                }
                string hex = value.Value.ToRgbHex();
                change ??= new A.ColorChange();
                change.RemoveAllChildren();
                change.ColorFrom = new A.ColorFrom(
                    new A.RgbColorModelHex { Val = hex });
                change.ColorTo = new A.ColorTo(
                    new A.RgbColorModelHex(new A.Alpha { Val = 0 }) {
                        Val = hex
                    });
                if (change.Parent == null) blip.Append(change);
            }
        }

        /// <summary>
        /// Gets or sets the RGB color used to recolor the entire picture, or
        /// <see langword="null"/> to disable recoloring.
        /// </summary>
        public OfficeColor? RecolorColor {
            get {
                A.ColorReplacement? replacement = Picture.BlipFill?.Blip?
                    .GetFirstChild<A.ColorReplacement>();
                A.RgbColorModelHex? rgb = replacement?
                    .GetFirstChild<A.RgbColorModelHex>();
                string? value = replacement?.ChildElements.Count == 1
                    && rgb?.ChildElements.Count == 0
                    ? rgb.Val?.Value
                    : null;
                return OfficeColor.TryParse(value, out OfficeColor color)
                    ? color
                    : null;
            }
            set {
                A.Blip blip = GetRequiredBlip();
                A.ColorReplacement? replacement = blip
                    .GetFirstChild<A.ColorReplacement>();
                if (!value.HasValue) {
                    replacement?.Remove();
                    return;
                }
                replacement ??= new A.ColorReplacement();
                replacement.RemoveAllChildren();
                replacement.Append(new A.RgbColorModelHex {
                    Val = value.Value.ToRgbHex()
                });
                if (replacement.Parent == null) blip.Append(replacement);
            }
        }

        /// <summary>
        /// Removes the luminance, grayscale, bi-level, transparent-color,
        /// and whole-picture recolor effects supported by binary PowerPoint.
        /// </summary>
        public void ResetClassicEffects() {
            A.Blip blip = GetRequiredBlip();
            blip.RemoveAllChildren<A.LuminanceEffect>();
            blip.RemoveAllChildren<A.Grayscale>();
            blip.RemoveAllChildren<A.BiLevel>();
            blip.RemoveAllChildren<A.ColorChange>();
            blip.RemoveAllChildren<A.ColorReplacement>();
        }

        private void UpdateLuminance(int? brightness, int? contrast) {
            A.Blip blip = GetRequiredBlip();
            A.LuminanceEffect? luminance = blip
                .GetFirstChild<A.LuminanceEffect>();
            if (!brightness.HasValue && !contrast.HasValue) {
                luminance?.Remove();
                return;
            }
            if (luminance == null) {
                luminance = new A.LuminanceEffect();
                blip.Append(luminance);
            }
            luminance.Brightness = brightness.HasValue
                ? brightness.Value * 1000
                : null;
            luminance.Contrast = contrast.HasValue
                ? contrast.Value * 1000
                : null;
        }

        private A.Blip GetRequiredBlip() => Picture.BlipFill?.Blip
            ?? throw new InvalidOperationException(
                "The picture has no DrawingML image reference.");

        private static void ValidateEffectPercent(int? value,
            string parameterName) {
            if (value is < -100 or > 100) {
                throw new ArgumentOutOfRangeException(parameterName,
                    "The picture adjustment must be between -100 and 100.");
            }
        }
    }
}
