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
            BackgroundColor = OfficeColor.ParseHex("FFFFFF"),
            SurfaceColor = OfficeColor.ParseHex("F7FAFC"),
            PanelColor = OfficeColor.ParseHex("FFFFFF"),
            PanelBorderColor = OfficeColor.ParseHex("D8E1E8"),
            PrimaryTextColor = OfficeColor.ParseHex("30343B"),
            SecondaryTextColor = OfficeColor.ParseHex("65717D"),
            MutedTextColor = OfficeColor.ParseHex("8A949E"),
            AccentColor = OfficeColor.ParseHex("0098C8"),
            AccentDarkColor = OfficeColor.ParseHex("006AA1"),
            AccentLightColor = OfficeColor.ParseHex("D9F4FA"),
            AccentContrastColor = OfficeColor.ParseHex("FFFFFF"),
            Accent2Color = OfficeColor.ParseHex("00B7C9"),
            Accent3Color = OfficeColor.ParseHex("6F6AA6"),
            WarningColor = OfficeColor.ParseHex("F4A100"),
            PaletteStyle = PowerPointPaletteStyle.Auto,
            HeadingFontFamily = "Poppins",
            BodyFontFamily = "Lato"
        };

        /// <summary>
        ///     Parses a six-digit RGB brand accent, optionally prefixed with '#', and creates readable defaults.
        /// </summary>
        public static PowerPointDesignTheme FromBrand(string accentColor, string? name = null,
            string headingFontFamily = "Poppins", string bodyFontFamily = "Lato") {
            OfficeColor accent = OfficeColor.ParseHex(NormalizeHex(accentColor, nameof(accentColor)));
            return new PowerPointDesignTheme {
                Name = name ?? "OfficeIMO Brand Theme",
                AccentColor = accent,
                AccentDarkColor = Shade(accent, -0.28),
                AccentLightColor = Shade(accent, 0.82),
                Accent2Color = Shade(accent, 0.18),
                Accent3Color = Shade(RotateAccent(accent), 0.1),
                WarningColor = OfficeColor.ParseHex("F4A100"),
                PaletteStyle = PowerPointPaletteStyle.Auto,
                HeadingFontFamily = headingFontFamily,
                BodyFontFamily = bodyFontFamily
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
                TypographyStyle = TypographyStyle,
                HeadingFontFamily = HeadingFontFamily,
                BodyFontFamily = BodyFontFamily
            };
        }

        /// <summary>
        ///     Creates a deterministic palette variation from literal accents, keeping the primary brand accent but changing supporting colors.
        /// </summary>
        public PowerPointDesignTheme WithVariation(string seed) {
            ArgumentException.ThrowIfNullOrWhiteSpace(seed);

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
        ///     Creates a copy with a specific typography strategy.
        /// </summary>
        public PowerPointDesignTheme WithTypographyStyle(PowerPointTypographyStyle typographyStyle, string seed) {
            ArgumentException.ThrowIfNullOrWhiteSpace(seed);

            PowerPointDesignTheme clone = Clone();
            clone.ApplyTypographyStyle(typographyStyle, seed);
            return clone;
        }

        /// <summary>
        ///     Creates a copy with a specific supporting palette strategy while preserving the primary brand accent.
        /// </summary>
        public PowerPointDesignTheme WithPaletteStyle(PowerPointPaletteStyle paletteStyle, string seed) {
            ArgumentException.ThrowIfNullOrWhiteSpace(seed);

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
                    clone.BackgroundColor = OfficeColor.ParseHex("FFFFFF");
                    clone.SurfaceColor = Shade(clone.Accent3Color, 0.91);
                    clone.PanelColor = OfficeColor.ParseHex("FFFFFF");
                    clone.PanelBorderColor = Shade(clone.Accent3Color, 0.68);
                    clone.PaletteStyle = PowerPointPaletteStyle.Auto;
                    clone.TypographyStyle = PowerPointTypographyStyle.ModernSans;
                    clone.HeadingFontFamily = "Aptos Display";
                    clone.BodyFontFamily = "Aptos";
                    break;
                case PowerPointDesignMood.Energetic:
                    clone.SurfaceColor = Shade(clone.WarningColor, 0.88);
                    clone.PanelBorderColor = Shade(clone.Accent2Color, 0.62);
                    clone.AccentLightColor = Shade(clone.AccentColor, 0.70);
                    clone.PaletteStyle = PowerPointPaletteStyle.Auto;
                    clone.TypographyStyle = PowerPointTypographyStyle.FriendlySans;
                    clone.HeadingFontFamily = "Poppins";
                    clone.BodyFontFamily = "Lato";
                    break;
                case PowerPointDesignMood.Minimal:
                    clone.SurfaceColor = OfficeColor.ParseHex("F5F7F9");
                    clone.PanelBorderColor = OfficeColor.ParseHex("DDE4EA");
                    clone.Accent2Color = Shade(clone.AccentColor, 0.34);
                    clone.Accent3Color = Shade(clone.AccentDarkColor, 0.50);
                    clone.WarningColor = clone.Accent2Color;
                    clone.PaletteStyle = PowerPointPaletteStyle.Auto;
                    clone.TypographyStyle = PowerPointTypographyStyle.ModernSans;
                    clone.HeadingFontFamily = "Aptos Display";
                    clone.BodyFontFamily = "Aptos";
                    break;
                default:
                    break;
            }

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
        public OfficeColor BackgroundColor { get; set; } = OfficeColor.ParseHex("FFFFFF");

        /// <summary>
        ///     Alternate surface color for subtle bands and washes.
        /// </summary>
        public OfficeColor SurfaceColor { get; set; } = OfficeColor.ParseHex("F7FAFC");

        /// <summary>
        ///     Card and panel fill color.
        /// </summary>
        public OfficeColor PanelColor { get; set; } = OfficeColor.ParseHex("FFFFFF");

        /// <summary>
        ///     Card and panel outline color.
        /// </summary>
        public OfficeColor PanelBorderColor { get; set; } = OfficeColor.ParseHex("D8E1E8");

        /// <summary>
        ///     Primary text color.
        /// </summary>
        public OfficeColor PrimaryTextColor { get; set; } = OfficeColor.ParseHex("30343B");

        /// <summary>
        ///     Secondary text color.
        /// </summary>
        public OfficeColor SecondaryTextColor { get; set; } = OfficeColor.ParseHex("65717D");

        /// <summary>
        ///     Muted caption and chrome text color.
        /// </summary>
        public OfficeColor MutedTextColor { get; set; } = OfficeColor.ParseHex("8A949E");

        /// <summary>
        ///     Main accent color.
        /// </summary>
        public OfficeColor AccentColor { get; set; } = OfficeColor.ParseHex("0098C8");

        /// <summary>
        ///     Dark accent color for section slides and strong bands.
        /// </summary>
        public OfficeColor AccentDarkColor { get; set; } = OfficeColor.ParseHex("006AA1");

        /// <summary>
        ///     Light accent color for soft backgrounds and dividers.
        /// </summary>
        public OfficeColor AccentLightColor { get; set; } = OfficeColor.ParseHex("D9F4FA");

        /// <summary>
        ///     Text color used on top of accent fills.
        /// </summary>
        public OfficeColor AccentContrastColor { get; set; } = OfficeColor.ParseHex("FFFFFF");

        /// <summary>
        ///     Secondary accent color.
        /// </summary>
        public OfficeColor Accent2Color { get; set; } = OfficeColor.ParseHex("00B7C9");

        /// <summary>
        ///     Tertiary accent color.
        /// </summary>
        public OfficeColor Accent3Color { get; set; } = OfficeColor.ParseHex("6F6AA6");

        /// <summary>
        ///     Warm accent used for markers and direction motifs.
        /// </summary>
        public OfficeColor WarningColor { get; set; } = OfficeColor.ParseHex("F4A100");

        /// <summary>
        ///     Supporting palette strategy used to generate secondary accents and surfaces.
        /// </summary>
        public PowerPointPaletteStyle PaletteStyle { get; private set; } = PowerPointPaletteStyle.Auto;

        /// <summary>
        ///     Typography strategy used to resolve the heading and body font pair.
        /// </summary>
        public PowerPointTypographyStyle TypographyStyle { get; private set; } = PowerPointTypographyStyle.Auto;

        /// <summary>
        ///     Font used for headings.
        /// </summary>
        public string HeadingFontFamily { get; set; } = "Poppins";

        /// <summary>
        ///     Font used for body text.
        /// </summary>
        public string BodyFontFamily { get; set; } = "Lato";

        internal void Validate() {

            ArgumentException.ThrowIfNullOrWhiteSpace(HeadingFontFamily);
            ArgumentException.ThrowIfNullOrWhiteSpace(BodyFontFamily);
        }

        internal void ApplyPaletteStyle(PowerPointPaletteStyle paletteStyle, string seed) {
            ArgumentException.ThrowIfNullOrWhiteSpace(seed);

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
                    SurfaceColor = OfficeColor.ParseHex("F8F5F0");
                    PanelBorderColor = OfficeColor.ParseHex("E4DDD2");
                    break;
                case PowerPointPaletteStyle.CoolNeutral:
                    Accent2Color = ShiftHue(AccentColor, -32d, saturationScale: 0.60, lightnessShift: 0.06);
                    Accent3Color = ShiftHue(AccentColor, 68d, saturationScale: 0.48, lightnessShift: 0.05);
                    WarningColor = WarmAccent(pick + 4);
                    SurfaceColor = OfficeColor.ParseHex("F4F7FA");
                    PanelBorderColor = OfficeColor.ParseHex("D9E2EA");
                    break;
                default:
                    break;
            }

            Validate();
        }

        internal void ApplyTypographyStyle(PowerPointTypographyStyle typographyStyle, string seed) {
            ArgumentException.ThrowIfNullOrWhiteSpace(seed);

            PowerPointTypographyStyle resolvedStyle = ResolveTypographyStyle(typographyStyle, seed);
            TypographyStyle = resolvedStyle;

            switch (resolvedStyle) {
                case PowerPointTypographyStyle.FriendlySans:
                    HeadingFontFamily = "Poppins";
                    BodyFontFamily = "Lato";
                    break;
                case PowerPointTypographyStyle.ModernSans:
                    HeadingFontFamily = "Aptos Display";
                    BodyFontFamily = "Aptos";
                    break;
                case PowerPointTypographyStyle.ExecutiveSans:
                    HeadingFontFamily = "Segoe UI Semibold";
                    BodyFontFamily = "Segoe UI";
                    break;
                case PowerPointTypographyStyle.EditorialSerif:
                    HeadingFontFamily = "Georgia";
                    BodyFontFamily = "Aptos";
                    break;
                case PowerPointTypographyStyle.TechnicalSans:
                    HeadingFontFamily = "Bahnschrift";
                    BodyFontFamily = "Aptos";
                    break;
                default:
                    break;
            }

            Validate();
        }

        internal void SyncTypographyStyleToFonts() {
            TypographyStyle = InferTypographyStyle(HeadingFontFamily, BodyFontFamily);
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

        private static OfficeColor Shade(OfficeColor value, double amount) {
            int r = value.R;
            int g = value.G;
            int b = value.B;
            r = ShadeChannel(r, amount);
            g = ShadeChannel(g, amount);
            b = ShadeChannel(b, amount);
            return OfficeColor.FromRgba((byte)r, (byte)g, (byte)b, value.A);
        }

        private static int ShadeChannel(int value, double amount) {
            double target = amount >= 0 ? 255 : 0;
            return (int)Math.Round(value + (target - value) * Math.Abs(amount));
        }

        private static OfficeColor RotateAccent(OfficeColor value) {
            return OfficeColor.FromRgba(value.G, value.B, value.R, value.A);
        }

        private static OfficeColor WarmAccent(int index) {
            string[] colors = {
                "F4A100",
                "F59E0B",
                "E9B44C",
                "FF8A3D",
                "D4AF37",
                "F97316",
                "EF4444"
            };
            return OfficeColor.ParseHex(colors[Math.Abs(index) % colors.Length]);
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

        private static PowerPointTypographyStyle ResolveTypographyStyle(PowerPointTypographyStyle typographyStyle,
            string seed) {
            if (typographyStyle != PowerPointTypographyStyle.Auto) {
                return typographyStyle;
            }

            PowerPointTypographyStyle[] styles = {
                PowerPointTypographyStyle.FriendlySans,
                PowerPointTypographyStyle.ModernSans,
                PowerPointTypographyStyle.ExecutiveSans,
                PowerPointTypographyStyle.EditorialSerif,
                PowerPointTypographyStyle.TechnicalSans
            };
            return styles[StablePick(seed + "/typography", styles.Length)];
        }

        private static PowerPointTypographyStyle InferTypographyStyle(string headingFontFamily, string bodyFontFamily) {
            if (MatchesTypographyPair(headingFontFamily, bodyFontFamily, "Poppins", "Lato")) {
                return PowerPointTypographyStyle.FriendlySans;
            }
            if (MatchesTypographyPair(headingFontFamily, bodyFontFamily, "Aptos Display", "Aptos")) {
                return PowerPointTypographyStyle.ModernSans;
            }
            if (MatchesTypographyPair(headingFontFamily, bodyFontFamily, "Segoe UI Semibold", "Segoe UI")) {
                return PowerPointTypographyStyle.ExecutiveSans;
            }
            if (MatchesTypographyPair(headingFontFamily, bodyFontFamily, "Georgia", "Aptos")) {
                return PowerPointTypographyStyle.EditorialSerif;
            }
            if (MatchesTypographyPair(headingFontFamily, bodyFontFamily, "Bahnschrift", "Aptos")) {
                return PowerPointTypographyStyle.TechnicalSans;
            }

            return PowerPointTypographyStyle.Auto;
        }

        private static bool MatchesTypographyPair(string headingFontFamily, string bodyFontFamily,
            string expectedHeadingFontFamily, string expectedBodyFontFamily) {
            return string.Equals(headingFontFamily, expectedHeadingFontFamily, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(bodyFontFamily, expectedBodyFontFamily, StringComparison.OrdinalIgnoreCase);
        }

        private static OfficeColor ShiftHue(OfficeColor value, double degrees, double saturationScale, double lightnessShift) {
            int r = value.R;
            int g = value.G;
            int b = value.B;
            RgbToHsl(r, g, b, out double hue, out double saturation, out double lightness);
            hue = (hue + degrees) % 360;
            if (hue < 0) {
                hue += 360;
            }
            saturation = Clamp(saturation * saturationScale, 0.28, 0.9);
            lightness = Clamp(lightness + lightnessShift, 0.26, 0.62);
            HslToRgb(hue, saturation, lightness, out r, out g, out b);
            return OfficeColor.FromRgba((byte)r, (byte)g, (byte)b, value.A);
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
