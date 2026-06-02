using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using Anchor = DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor;
using ShapeProperties = DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties;
using V = DocumentFormat.OpenXml.Vml;

#nullable enable annotations
using DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents an image contained in a <see cref="WordDocument"/> and provides
    /// functionality to insert and manipulate pictures.
    /// </summary>
    public partial class WordImage : WordElement {

        /// <summary>
        /// Sets a fixed opacity value for the image.
        /// </summary>
        public int? FixedOpacity {
            get {
                var blip = GetBlip();
                var ar = blip?.GetFirstChild<AlphaReplace>();
                var alpha = ar?.Alpha?.Value;
                _fixedOpacity = alpha != null ? (int?)(alpha / 1000) : null;
                return _fixedOpacity;
            }
            set {
                if (value is < 0 or > 100)
                    throw new ArgumentOutOfRangeException(nameof(value), "Opacity must be between 0 and 100.");

                _fixedOpacity = value;
                var blip = GetBlip();
                if (blip == null) return;
                var ar = blip.GetFirstChild<AlphaReplace>();
                if (value == null) {
                    ar?.Remove();
                    return;
                }
                if (ar == null) {
                    ar = new AlphaReplace();
                    blip.Append(ar);
                }
                ar.Alpha = value.Value * 1000;
            }
        }

        /// <summary>
        /// Gets or sets the alpha inversion color in hex.
        /// </summary>
        public string? AlphaInversionColorHex {
            get {
                var blip = GetBlip();
                var ai = blip?.GetFirstChild<AlphaInverse>();
                _alphaInversionColorHex = ai?.GetFirstChild<RgbColorModelHex>()?.Val;
                return _alphaInversionColorHex;
            }
            set {
                _alphaInversionColorHex = value;
                var blip = GetBlip();
                if (blip == null) return;
                var ai = blip.GetFirstChild<AlphaInverse>();
                if (value == null) {
                    ai?.Remove();
                    return;
                }
                if (ai == null) {
                    ai = new AlphaInverse();
                    blip.Append(ai);
                }
                var clr = ai.GetFirstChild<RgbColorModelHex>();
                if (clr == null) {
                    ai.RemoveAllChildren();
                    clr = new RgbColorModelHex();
                    ai.Append(clr);
                }
                clr.Val = value;
            }
        }

        /// <summary>
        /// Gets or sets the alpha inversion color.
        /// </summary>
        public OfficeIMO.Drawing.OfficeColor? AlphaInversionColor {
            get {
                if (AlphaInversionColorHex == null) return (OfficeIMO.Drawing.OfficeColor?)null;
                return Helpers.ParseColor(AlphaInversionColorHex);
            }
            set { AlphaInversionColorHex = value?.ToHexColor(); }
        }

        /// <summary>
        /// Gets or sets the threshold for black and white effect.
        /// </summary>
        public int? BlackWhiteThreshold {
            get {
                var blip = GetBlip();
                var bi = blip?.GetFirstChild<BiLevel>();
                var threshold = bi?.Threshold?.Value;
                _blackWhiteThreshold = threshold != null ? (int?)(threshold / 1000) : null;
                return _blackWhiteThreshold;
            }
            set {
                if (value is < 0 or > 100)
                    throw new ArgumentOutOfRangeException(nameof(value), "Threshold must be between 0 and 100.");
                _blackWhiteThreshold = value;
                var blip = GetBlip();
                if (blip == null) return;
                var bi = blip.GetFirstChild<BiLevel>();
                if (value == null) {
                    bi?.Remove();
                    return;
                }
                if (bi == null) {
                    bi = new BiLevel();
                    blip.Append(bi);
                }
                bi.Threshold = value.Value * 1000;
            }
        }

        /// <summary>
        /// Gets or sets the blur radius in EMUs.
        /// </summary>
        public int? BlurRadius {
            get {
                var blip = GetBlip();
                var blur = blip?.GetFirstChild<Blur>();
                var radius = blur?.Radius?.Value;
                _blurRadius = radius != null ? (int?)radius : null;
                return _blurRadius;
            }
            set {
                _blurRadius = value;
                var blip = GetBlip();
                if (blip == null) return;
                var blur = blip.GetFirstChild<Blur>();
                if (value == null && _blurGrow == null) {
                    blur?.Remove();
                    return;
                }
                if (blur == null) {
                    blur = new Blur();
                    blip.Append(blur);
                }
                blur.Radius = value != null ? new Int64Value((long)value.Value) : null;
                blur.Grow = _blurGrow ?? false;
            }
        }

        /// <summary>
        /// Gets or sets whether the blur grows the image's bounds.
        /// </summary>
        public bool? BlurGrow {
            get {
                var blip = GetBlip();
                var blur = blip?.GetFirstChild<Blur>();
                _blurGrow = blur?.Grow?.Value;
                return _blurGrow;
            }
            set {
                _blurGrow = value;
                var blip = GetBlip();
                if (blip == null) return;
                var blur = blip.GetFirstChild<Blur>();
                if (value == null && _blurRadius == null) {
                    blur?.Remove();
                    return;
                }
                if (blur == null) {
                    blur = new Blur();
                    blip.Append(blur);
                }
                blur.Radius = _blurRadius != null ? new Int64Value((long)_blurRadius.Value) : null;
                blur.Grow = value ?? false;
            }
        }

        /// <summary>
        /// Gets or sets the source color to change from in hex.
        /// </summary>
        public string? ColorChangeFromHex {
            get {
                var blip = GetBlip();
                var cc = blip?.GetFirstChild<ColorChange>();
                _colorChangeFromHex = cc?.ColorFrom?.GetFirstChild<RgbColorModelHex>()?.Val;
                return _colorChangeFromHex;
            }
            set {
                _colorChangeFromHex = value;
                UpdateColorChange();
            }
        }

        /// <summary>
        /// Gets or sets the target color to change to in hex.
        /// </summary>
        public string? ColorChangeToHex {
            get {
                var blip = GetBlip();
                var cc = blip?.GetFirstChild<ColorChange>();
                _colorChangeToHex = cc?.ColorTo?.GetFirstChild<RgbColorModelHex>()?.Val;
                return _colorChangeToHex;
            }
            set {
                _colorChangeToHex = value;
                UpdateColorChange();
            }
        }

        /// <summary>
        /// Gets or sets the source color to change from.
        /// </summary>
        public OfficeIMO.Drawing.OfficeColor? ColorChangeFrom {
            get {
                return ColorChangeFromHex == null ? (OfficeIMO.Drawing.OfficeColor?)null : Helpers.ParseColor(ColorChangeFromHex);
            }
            set { ColorChangeFromHex = value?.ToHexColor(); }
        }

        /// <summary>
        /// Gets or sets the target color to change to.
        /// </summary>
        public OfficeIMO.Drawing.OfficeColor? ColorChangeTo {
            get {
                return ColorChangeToHex == null ? (OfficeIMO.Drawing.OfficeColor?)null : Helpers.ParseColor(ColorChangeToHex);
            }
            set { ColorChangeToHex = value?.ToHexColor(); }
        }

        private void UpdateColorChange() {
            var blip = GetBlip();
            if (blip == null) return;
            var cc = blip.GetFirstChild<ColorChange>();
            if (_colorChangeFromHex == null && _colorChangeToHex == null) {
                cc?.Remove();
                return;
            }
            if (cc == null) {
                cc = new ColorChange();
                blip.Append(cc);
            }
            if (_colorChangeFromHex != null) {
                var from = cc.ColorFrom ?? new ColorFrom();
                var clr = from.GetFirstChild<RgbColorModelHex>() ?? new RgbColorModelHex();
                clr.Val = _colorChangeFromHex;
                if (from.FirstChild == null) from.Append(clr);
                cc.ColorFrom = from;
            }
            if (_colorChangeToHex != null) {
                var to = cc.ColorTo ?? new ColorTo();
                var clr = to.GetFirstChild<RgbColorModelHex>() ?? new RgbColorModelHex();
                clr.Val = _colorChangeToHex;
                if (to.FirstChild == null) to.Append(clr);
                cc.ColorTo = to;
            }
        }

        /// <summary>
        /// Gets or sets a color replacement hex value.
        /// </summary>
        public string? ColorReplacementHex {
            get {
                var blip = GetBlip();
                var cr = blip?.GetFirstChild<ColorReplacement>();
                _colorReplacementHex = cr?.GetFirstChild<RgbColorModelHex>()?.Val;
                return _colorReplacementHex;
            }
            set {
                _colorReplacementHex = value;
                var blip = GetBlip();
                if (blip == null) return;
                var cr = blip.GetFirstChild<ColorReplacement>();
                if (value == null) {
                    cr?.Remove();
                    return;
                }
                if (cr == null) {
                    cr = new ColorReplacement();
                    blip.Append(cr);
                }
                var clr = cr.GetFirstChild<RgbColorModelHex>();
                if (clr == null) {
                    cr.RemoveAllChildren();
                    clr = new RgbColorModelHex();
                    cr.Append(clr);
                }
                clr.Val = value;
            }
        }

        /// <summary>
        /// Gets or sets a color replacement.
        /// </summary>
        public OfficeIMO.Drawing.OfficeColor? ColorReplacement {
            get {
                return ColorReplacementHex == null ? (OfficeIMO.Drawing.OfficeColor?)null : Helpers.ParseColor(ColorReplacementHex);
            }
            set { ColorReplacementHex = value?.ToHexColor(); }
        }

        /// <summary>
        /// Gets or sets the first duotone color in hex.
        /// </summary>
        public string? DuotoneColor1Hex {
            get {
                var blip = GetBlip();
                var duo = blip?.GetFirstChild<Duotone>();
                _duotoneColor1Hex = duo?.GetFirstChild<RgbColorModelHex>()?.Val;
                return _duotoneColor1Hex;
            }
            set {
                _duotoneColor1Hex = value;
                UpdateDuotone();
            }
        }

        /// <summary>
        /// Gets or sets the second duotone color in hex.
        /// </summary>
        public string? DuotoneColor2Hex {
            get {
                var blip = GetBlip();
                var duo = blip?.GetFirstChild<Duotone>();
                _duotoneColor2Hex = duo?.Elements<RgbColorModelHex>().Skip(1).FirstOrDefault()?.Val;
                return _duotoneColor2Hex;
            }
            set {
                _duotoneColor2Hex = value;
                UpdateDuotone();
            }
        }

        /// <summary>
        /// Gets or sets the first duotone color.
        /// </summary>
        public OfficeIMO.Drawing.OfficeColor? DuotoneColor1 {
            get {
                return DuotoneColor1Hex == null ? (OfficeIMO.Drawing.OfficeColor?)null : Helpers.ParseColor(DuotoneColor1Hex);
            }
            set { DuotoneColor1Hex = value?.ToHexColor(); }
        }

        /// <summary>
        /// Gets or sets the second duotone color.
        /// </summary>
        public OfficeIMO.Drawing.OfficeColor? DuotoneColor2 {
            get {
                return DuotoneColor2Hex == null ? (OfficeIMO.Drawing.OfficeColor?)null : Helpers.ParseColor(DuotoneColor2Hex);
            }
            set { DuotoneColor2Hex = value?.ToHexColor(); }
        }

        private void UpdateDuotone() {
            var blip = GetBlip();
            if (blip == null) return;
            var duo = blip.GetFirstChild<Duotone>();
            if (_duotoneColor1Hex == null && _duotoneColor2Hex == null) {
                duo?.Remove();
                return;
            }
            if (duo == null) {
                duo = new Duotone();
                blip.Append(duo);
            }
            duo.RemoveAllChildren();
            if (_duotoneColor1Hex != null)
                duo.Append(new RgbColorModelHex { Val = _duotoneColor1Hex });
            if (_duotoneColor2Hex != null)
                duo.Append(new RgbColorModelHex { Val = _duotoneColor2Hex });
        }

        /// <summary>
        /// Gets or sets whether the image is displayed in grayscale.
        /// </summary>
        public bool? GrayScale {
            get {
                var blip = GetBlip();
                var gs = blip?.GetFirstChild<Grayscale>();
                _grayScale = gs != null;
                return _grayScale;
            }
            set {
                _grayScale = value;
                var blip = GetBlip();
                if (blip == null) return;
                var gs = blip.GetFirstChild<Grayscale>();
                if (value == true && gs == null) {
                    blip.Append(new Grayscale());
                } else if (value != true) {
                    gs?.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets the brightness adjustment.
        /// </summary>
        public int? LuminanceBrightness {
            get {
                var blip = GetBlip();
                var lum = blip?.GetFirstChild<LuminanceEffect>();
                var brightness = lum?.Brightness?.Value;
                _luminanceBrightness = brightness != null ? (int?)brightness : null;
                if (_luminanceBrightness != null) _luminanceBrightness /= 1000;
                return _luminanceBrightness;
            }
            set {
                _luminanceBrightness = value;
                UpdateLuminance();
            }
        }

        /// <summary>
        /// Gets or sets the contrast adjustment.
        /// </summary>
        public int? LuminanceContrast {
            get {
                var blip = GetBlip();
                var lum = blip?.GetFirstChild<LuminanceEffect>();
                _luminanceContrast = lum?.Contrast?.Value;
                if (_luminanceContrast != null) _luminanceContrast /= 1000;
                return _luminanceContrast;
            }
            set {
                _luminanceContrast = value;
                UpdateLuminance();
            }
        }

        private void UpdateLuminance() {
            var blip = GetBlip();
            if (blip == null) return;
            var lum = blip.GetFirstChild<LuminanceEffect>();
            if (_luminanceBrightness == null && _luminanceContrast == null) {
                lum?.Remove();
                return;
            }
            if (lum == null) {
                lum = new LuminanceEffect();
                blip.Append(lum);
            }
            lum.Brightness = _luminanceBrightness != null ? new Int32Value(_luminanceBrightness.Value * 1000) : null;
            lum.Contrast = _luminanceContrast != null ? new Int32Value(_luminanceContrast.Value * 1000) : null;
        }

        /// <summary>
        /// Gets or sets the tint amount.
        /// </summary>
        public int? TintAmount {
            get {
                var blip = GetBlip();
                var tint = blip?.GetFirstChild<TintEffect>();
                _tintAmount = tint?.Amount != null ? tint.Amount / 1000 : null;
                return _tintAmount;
            }
            set {
                _tintAmount = value;
                UpdateTint();
            }
        }

        /// <summary>
        /// Gets or sets the tint hue.
        /// </summary>
        public int? TintHue {
            get {
                var blip = GetBlip();
                var tint = blip?.GetFirstChild<TintEffect>();
                _tintHue = tint?.Hue != null ? tint.Hue / 60000 : null;
                return _tintHue;
            }
            set {
                _tintHue = value;
                UpdateTint();
            }
        }

        private void UpdateTint() {
            var blip = GetBlip();
            if (blip == null) return;
            var tint = blip.GetFirstChild<TintEffect>();
            if (_tintAmount == null && _tintHue == null) {
                tint?.Remove();
                return;
            }
            if (tint == null) {
                tint = new TintEffect();
                blip.Append(tint);
            }
            tint.Amount = _tintAmount != null ? new Int32Value(_tintAmount.Value * 1000) : null;
            tint.Hue = _tintHue != null ? new Int32Value(_tintHue.Value * 60000) : null;
        }

    }
}
