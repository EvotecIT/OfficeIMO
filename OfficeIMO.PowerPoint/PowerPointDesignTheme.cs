using System;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Defines a reusable visual system for high-level PowerPoint slide compositions.
    /// </summary>
    public sealed class PowerPointDesignTheme {
        /// <summary>
        ///     Creates a modern blue theme suitable for clean business decks.
        /// </summary>
        public static PowerPointDesignTheme ModernBlue => new() {
            Name = "OfficeIMO Modern Blue",
            BackgroundColor = "FFFFFF",
            SurfaceColor = "F7FAFC",
            PanelColor = "FFFFFF",
            PanelBorderColor = "D8E1E8",
            PrimaryTextColor = "30343B",
            SecondaryTextColor = "65717D",
            MutedTextColor = "8A949E",
            AccentColor = "0098C8",
            AccentDarkColor = "006AA1",
            AccentLightColor = "D9F4FA",
            AccentContrastColor = "FFFFFF",
            Accent2Color = "00B7C9",
            Accent3Color = "6F6AA6",
            WarningColor = "F4A100",
            PaletteStyle = PowerPointPaletteStyle.Auto,
            HeadingFontName = "Poppins",
            BodyFontName = "Lato"
        };

        /// <summary>
        ///     Creates a theme from a primary brand accent while keeping readable defaults.
        /// </summary>
        public static PowerPointDesignTheme FromBrand(string accentColor, string? name = null,
            string headingFontName = "Poppins", string bodyFontName = "Lato") {
            string accent = NormalizeHex(accentColor, nameof(accentColor));
            return new PowerPointDesignTheme {
                Name = name ?? "OfficeIMO Brand Theme",
                AccentColor = accent,
                AccentDarkColor = Shade(accent, -0.28),
                AccentLightColor = Shade(accent, 0.82),
                Accent2Color = Shade(accent, 0.18),
                Accent3Color = Shade(RotateAccent(accent), 0.1),
                WarningColor = "F4A100",
                PaletteStyle = PowerPointPaletteStyle.Auto,
                HeadingFontName = headingFontName,
                BodyFontName = bodyFontName
            };
        }

        /// <summary>
        ///     Creates a copy of the theme that can be modified without changing the source instance.
        /// </summary>
        public PowerPointDesignTheme Clone() {
            return new PowerPointDesignTheme {
                Name = Name,
                BackgroundColor = BackgroundColor,
                SurfaceColor = SurfaceColor,
                PanelColor = PanelColor,
                PanelBorderColor = PanelBorderColor,
                PrimaryTextColor = PrimaryTextColor,
                SecondaryTextColor = SecondaryTextColor,
                MutedTextColor = MutedTextColor,
                AccentColor = AccentColor,
                AccentDarkColor = AccentDarkColor,
                AccentLightColor = AccentLightColor,
                AccentContrastColor = AccentContrastColor,
                Accent2Color = Accent2Color,
                Accent3Color = Accent3Color,
                WarningColor = WarningColor,
                PaletteStyle = PaletteStyle,
                HeadingFontName = HeadingFontName,
                BodyFontName = BodyFontName
            };
        }

        /// <summary>
        ///     Creates a deterministic palette variation that keeps the primary brand accent but changes supporting colors.
        /// </summary>
        public PowerPointDesignTheme WithVariation(string seed) {
            if (string.IsNullOrWhiteSpace(seed)) {
                throw new ArgumentException("Variation seed cannot be null or empty.", nameof(seed));
            }

            int pick = StablePick(seed, 7);
            double coolShift = 18d + pick * 8d;

            PowerPointDesignTheme clone = Clone();
            clone.Name = Name + " Variant " + (pick + 1);
            clone.Accent2Color = ShiftHue(AccentColor, coolShift, saturationScale: 0.95, lightnessShift: 0.02);
            clone.Accent3Color = ShiftHue(AccentColor, -coolShift - 18, saturationScale: 0.75, lightnessShift: 0.04);
            clone.WarningColor = WarmAccent(pick);
            clone.AccentLightColor = Shade(AccentColor, 0.78 + pick * 0.02);
            clone.SurfaceColor = Shade(clone.Accent2Color, 0.91);
            clone.PanelBorderColor = Shade(clone.Accent3Color, 0.72);
            clone.PaletteStyle = PowerPointPaletteStyle.Auto;
            clone.Validate();
            return clone;
        }

        /// <summary>
        ///     Creates a copy with a specific supporting palette strategy while preserving the primary brand accent.
        /// </summary>
        public PowerPointDesignTheme WithPaletteStyle(PowerPointPaletteStyle paletteStyle, string seed) {
            if (string.IsNullOrWhiteSpace(seed)) {
                throw new ArgumentException("Palette seed cannot be null or empty.", nameof(seed));
            }

            PowerPointDesignTheme clone = Clone();
            clone.ApplyPaletteStyle(paletteStyle, seed);
            return clone;
        }

        /// <summary>
        ///     Creates a theme variation tuned for a broad deck mood.
        /// </summary>
        public PowerPointDesignTheme WithMood(PowerPointDesignMood mood) {
            PowerPointDesignTheme clone = Clone();
            clone.Name = Name + " " + mood;

            switch (mood) {
                case PowerPointDesignMood.Editorial:
                    clone.BackgroundColor = "FFFFFF";
                    clone.SurfaceColor = Shade(clone.Accent3Color, 0.91);
                    clone.PanelColor = "FFFFFF";
                    clone.PanelBorderColor = Shade(clone.Accent3Color, 0.68);
                    clone.HeadingFontName = "Aptos Display";
                    clone.BodyFontName = "Aptos";
                    break;
                case PowerPointDesignMood.Energetic:
                    clone.SurfaceColor = Shade(clone.WarningColor, 0.88);
                    clone.PanelBorderColor = Shade(clone.Accent2Color, 0.62);
                    clone.AccentLightColor = Shade(clone.AccentColor, 0.70);
                    clone.HeadingFontName = "Poppins";
                    clone.BodyFontName = "Aptos";
                    break;
                case PowerPointDesignMood.Minimal:
                    clone.SurfaceColor = "F5F7F9";
                    clone.PanelBorderColor = "DDE4EA";
                    clone.Accent2Color = Shade(clone.AccentColor, 0.34);
                    clone.Accent3Color = Shade(clone.AccentDarkColor, 0.50);
                    clone.WarningColor = clone.Accent2Color;
                    clone.HeadingFontName = "Aptos Display";
                    clone.BodyFontName = "Aptos";
                    break;
                default:
                    break;
            }

            clone.PaletteStyle = PowerPointPaletteStyle.Auto;
            clone.Validate();
            return clone;
        }

        /// <summary>
        ///     Theme display name.
        /// </summary>
        public string Name { get; set; } = "OfficeIMO Design";

        /// <summary>
        ///     Main slide background color.
        /// </summary>
        public string BackgroundColor { get; set; } = "FFFFFF";

        /// <summary>
        ///     Alternate surface color for subtle bands and washes.
        /// </summary>
        public string SurfaceColor { get; set; } = "F7FAFC";

        /// <summary>
        ///     Card and panel fill color.
        /// </summary>
        public string PanelColor { get; set; } = "FFFFFF";

        /// <summary>
        ///     Card and panel outline color.
        /// </summary>
        public string PanelBorderColor { get; set; } = "D8E1E8";

        /// <summary>
        ///     Primary text color.
        /// </summary>
        public string PrimaryTextColor { get; set; } = "30343B";

        /// <summary>
        ///     Secondary text color.
        /// </summary>
        public string SecondaryTextColor { get; set; } = "65717D";

        /// <summary>
        ///     Muted caption and chrome text color.
        /// </summary>
        public string MutedTextColor { get; set; } = "8A949E";

        /// <summary>
        ///     Main accent color.
        /// </summary>
        public string AccentColor { get; set; } = "0098C8";

        /// <summary>
        ///     Dark accent color for section slides and strong bands.
        /// </summary>
        public string AccentDarkColor { get; set; } = "006AA1";

        /// <summary>
        ///     Light accent color for soft backgrounds and dividers.
        /// </summary>
        public string AccentLightColor { get; set; } = "D9F4FA";

        /// <summary>
        ///     Text color used on top of accent fills.
        /// </summary>
        public string AccentContrastColor { get; set; } = "FFFFFF";

        /// <summary>
        ///     Secondary accent color.
        /// </summary>
        public string Accent2Color { get; set; } = "00B7C9";

        /// <summary>
        ///     Tertiary accent color.
        /// </summary>
        public string Accent3Color { get; set; } = "6F6AA6";

        /// <summary>
        ///     Warm accent used for markers and direction motifs.
        /// </summary>
        public string WarningColor { get; set; } = "F4A100";

        /// <summary>
        ///     Supporting palette strategy used to generate secondary accents and surfaces.
        /// </summary>
        public PowerPointPaletteStyle PaletteStyle { get; private set; } = PowerPointPaletteStyle.Auto;

        /// <summary>
        ///     Font used for headings.
        /// </summary>
        public string HeadingFontName { get; set; } = "Poppins";

        /// <summary>
        ///     Font used for body text.
        /// </summary>
        public string BodyFontName { get; set; } = "Lato";

        internal void Validate() {
            ValidateHex(BackgroundColor, nameof(BackgroundColor));
            ValidateHex(SurfaceColor, nameof(SurfaceColor));
            ValidateHex(PanelColor, nameof(PanelColor));
            ValidateHex(PanelBorderColor, nameof(PanelBorderColor));
            ValidateHex(PrimaryTextColor, nameof(PrimaryTextColor));
            ValidateHex(SecondaryTextColor, nameof(SecondaryTextColor));
            ValidateHex(MutedTextColor, nameof(MutedTextColor));
            ValidateHex(AccentColor, nameof(AccentColor));
            ValidateHex(AccentDarkColor, nameof(AccentDarkColor));
            ValidateHex(AccentLightColor, nameof(AccentLightColor));
            ValidateHex(AccentContrastColor, nameof(AccentContrastColor));
            ValidateHex(Accent2Color, nameof(Accent2Color));
            ValidateHex(Accent3Color, nameof(Accent3Color));
            ValidateHex(WarningColor, nameof(WarningColor));

            if (string.IsNullOrWhiteSpace(HeadingFontName)) {
                throw new ArgumentException("Heading font cannot be null or empty.", nameof(HeadingFontName));
            }
            if (string.IsNullOrWhiteSpace(BodyFontName)) {
                throw new ArgumentException("Body font cannot be null or empty.", nameof(BodyFontName));
            }
        }

        internal void ApplyPaletteStyle(PowerPointPaletteStyle paletteStyle, string seed) {
            if (string.IsNullOrWhiteSpace(seed)) {
                throw new ArgumentException("Palette seed cannot be null or empty.", nameof(seed));
            }

            PowerPointPaletteStyle resolvedStyle = ResolvePaletteStyle(paletteStyle, seed);
            int pick = StablePick(seed + "/" + resolvedStyle, 7);
            PaletteStyle = resolvedStyle;

            switch (resolvedStyle) {
                case PowerPointPaletteStyle.Analogous:
                    Accent2Color = ShiftHue(AccentColor, 18d + pick * 3d, saturationScale: 0.92,
                        lightnessShift: 0.03);
                    Accent3Color = ShiftHue(AccentColor, -28d - pick * 3d, saturationScale: 0.78,
                        lightnessShift: 0.04);
                    WarningColor = WarmAccent(pick);
                    SurfaceColor = Shade(Accent2Color, 0.90);
                    PanelBorderColor = Shade(Accent3Color, 0.70);
                    break;
                case PowerPointPaletteStyle.Complementary:
                    Accent2Color = ShiftHue(AccentColor, 174d + pick, saturationScale: 0.72,
                        lightnessShift: 0.02);
                    Accent3Color = ShiftHue(AccentColor, 42d + pick * 2d, saturationScale: 0.82,
                        lightnessShift: 0.05);
                    WarningColor = WarmAccent(pick + 1);
                    SurfaceColor = Shade(Accent2Color, 0.91);
                    PanelBorderColor = Shade(Accent2Color, 0.68);
                    break;
                case PowerPointPaletteStyle.SplitComplementary:
                    Accent2Color = ShiftHue(AccentColor, 146d + pick * 2d, saturationScale: 0.74,
                        lightnessShift: 0.02);
                    Accent3Color = ShiftHue(AccentColor, -148d - pick * 2d, saturationScale: 0.78,
                        lightnessShift: 0.03);
                    WarningColor = WarmAccent(pick + 2);
                    SurfaceColor = Shade(Accent3Color, 0.91);
                    PanelBorderColor = Shade(Accent2Color, 0.68);
                    break;
                case PowerPointPaletteStyle.Monochrome:
                    Accent2Color = Shade(AccentColor, 0.28);
                    Accent3Color = Shade(AccentDarkColor, 0.45);
                    WarningColor = Shade(AccentColor, 0.52);
                    SurfaceColor = Shade(AccentColor, 0.93);
                    PanelBorderColor = Shade(AccentColor, 0.72);
                    break;
                case PowerPointPaletteStyle.WarmNeutral:
                    Accent2Color = WarmAccent(pick);
                    Accent3Color = ShiftHue(AccentColor, -38d, saturationScale: 0.62, lightnessShift: 0.03);
                    WarningColor = WarmAccent(pick + 3);
                    SurfaceColor = "F8F5F0";
                    PanelBorderColor = "E4DDD2";
                    break;
                case PowerPointPaletteStyle.CoolNeutral:
                    Accent2Color = ShiftHue(AccentColor, -32d, saturationScale: 0.60, lightnessShift: 0.06);
                    Accent3Color = ShiftHue(AccentColor, 68d, saturationScale: 0.48, lightnessShift: 0.05);
                    WarningColor = WarmAccent(pick + 4);
                    SurfaceColor = "F4F7FA";
                    PanelBorderColor = "D9E2EA";
                    break;
                default:
                    break;
            }

            Validate();
        }

        private static string NormalizeHex(string value, string name) {
            if (string.IsNullOrWhiteSpace(value)) {
                throw new ArgumentException("Color cannot be null or empty.", name);
            }
            if (value.StartsWith("#", StringComparison.Ordinal)) {
                value = value.Substring(1);
            }
            ValidateHex(value, name);
            return value.ToUpperInvariant();
        }

        private static string Shade(string value, double amount) {
            int r = Convert.ToInt32(value.Substring(0, 2), 16);
            int g = Convert.ToInt32(value.Substring(2, 2), 16);
            int b = Convert.ToInt32(value.Substring(4, 2), 16);
            r = ShadeChannel(r, amount);
            g = ShadeChannel(g, amount);
            b = ShadeChannel(b, amount);
            return r.ToString("X2") + g.ToString("X2") + b.ToString("X2");
        }

        private static int ShadeChannel(int value, double amount) {
            double target = amount >= 0 ? 255 : 0;
            return (int)Math.Round(value + (target - value) * Math.Abs(amount));
        }

        private static string RotateAccent(string value) {
            return value.Substring(2, 2) + value.Substring(4, 2) + value.Substring(0, 2);
        }

        private static string WarmAccent(int index) {
            string[] colors = {
                "F4A100",
                "F59E0B",
                "E9B44C",
                "FF8A3D",
                "D4AF37",
                "F97316",
                "EF4444"
            };
            return colors[Math.Abs(index) % colors.Length];
        }

        private static int StablePick(string value, int choices) {
            unchecked {
                int hash = (int)2166136261;
                for (int i = 0; i < value.Length; i++) {
                    hash ^= value[i];
                    hash *= 16777619;
                }
                return (hash & int.MaxValue) % choices;
            }
        }

        private static PowerPointPaletteStyle ResolvePaletteStyle(PowerPointPaletteStyle paletteStyle, string seed) {
            if (paletteStyle != PowerPointPaletteStyle.Auto) {
                return paletteStyle;
            }

            PowerPointPaletteStyle[] styles = {
                PowerPointPaletteStyle.Analogous,
                PowerPointPaletteStyle.Complementary,
                PowerPointPaletteStyle.SplitComplementary,
                PowerPointPaletteStyle.Monochrome,
                PowerPointPaletteStyle.WarmNeutral,
                PowerPointPaletteStyle.CoolNeutral
            };
            return styles[StablePick(seed + "/palette", styles.Length)];
        }

        private static string ShiftHue(string value, double degrees, double saturationScale, double lightnessShift) {
            int r = Convert.ToInt32(value.Substring(0, 2), 16);
            int g = Convert.ToInt32(value.Substring(2, 2), 16);
            int b = Convert.ToInt32(value.Substring(4, 2), 16);
            RgbToHsl(r, g, b, out double hue, out double saturation, out double lightness);
            hue = (hue + degrees) % 360;
            if (hue < 0) {
                hue += 360;
            }
            saturation = Clamp(saturation * saturationScale, 0.28, 0.9);
            lightness = Clamp(lightness + lightnessShift, 0.26, 0.62);
            HslToRgb(hue, saturation, lightness, out r, out g, out b);
            return r.ToString("X2") + g.ToString("X2") + b.ToString("X2");
        }

        private static void RgbToHsl(int r, int g, int b, out double hue, out double saturation, out double lightness) {
            double red = r / 255d;
            double green = g / 255d;
            double blue = b / 255d;
            double max = Math.Max(red, Math.Max(green, blue));
            double min = Math.Min(red, Math.Min(green, blue));
            double delta = max - min;

            lightness = (max + min) / 2d;
            const double epsilon = 1e-12;
            if (Math.Abs(delta) < epsilon) {
                hue = 0;
                saturation = 0;
                return;
            }

            saturation = delta / (1d - Math.Abs(2d * lightness - 1d));
            if (red >= green && red >= blue) {
                hue = 60d * (((green - blue) / delta) % 6d);
            } else if (green >= red && green >= blue) {
                hue = 60d * (((blue - red) / delta) + 2d);
            } else {
                hue = 60d * (((red - green) / delta) + 4d);
            }
            if (hue < 0) {
                hue += 360d;
            }
        }

        private static void HslToRgb(double hue, double saturation, double lightness, out int r, out int g, out int b) {
            double chroma = (1d - Math.Abs(2d * lightness - 1d)) * saturation;
            double x = chroma * (1d - Math.Abs((hue / 60d) % 2d - 1d));
            double m = lightness - chroma / 2d;
            double red;
            double green;
            double blue;

            if (hue < 60) {
                red = chroma;
                green = x;
                blue = 0;
            } else if (hue < 120) {
                red = x;
                green = chroma;
                blue = 0;
            } else if (hue < 180) {
                red = 0;
                green = chroma;
                blue = x;
            } else if (hue < 240) {
                red = 0;
                green = x;
                blue = chroma;
            } else if (hue < 300) {
                red = x;
                green = 0;
                blue = chroma;
            } else {
                red = chroma;
                green = 0;
                blue = x;
            }

            r = ToByte(red + m);
            g = ToByte(green + m);
            b = ToByte(blue + m);
        }

        private static int ToByte(double value) {
            return (int)Math.Round(Clamp(value, 0, 1) * 255d);
        }

        private static double Clamp(double value, double min, double max) {
            if (value < min) {
                return min;
            }
            if (value > max) {
                return max;
            }
            return value;
        }

        private static void ValidateHex(string value, string name) {
            if (string.IsNullOrWhiteSpace(value)) {
                throw new ArgumentException("Color cannot be null or empty.", name);
            }
            if (value.Length != 6) {
                throw new ArgumentException("Color must be a six-character RGB hex value without '#'.", name);
            }
            for (int i = 0; i < value.Length; i++) {
                char c = value[i];
                bool valid = c is >= '0' and <= '9' or >= 'A' and <= 'F' or >= 'a' and <= 'f';
                if (!valid) {
                    throw new ArgumentException("Color must be a six-character RGB hex value without '#'.", name);
                }
            }
        }
    }
}
