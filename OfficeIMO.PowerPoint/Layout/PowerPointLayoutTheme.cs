using System;
using System.Collections.Generic;
using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>Text styles of the slide layout system (Build/Visual/powerpoint-layouts.json, textStyles).</summary>
    internal enum PowerPointTextRole {
        Display,
        SectionTitle,
        SectionNumeral,
        Title,
        Subtitle,
        Lede,
        Eyebrow,
        Heading,
        Body,
        BodySmall,
        KpiValue,
        KpiLabel,
        Quote,
        TableHeader,
        TableBody,
        Caption,
        Footer
    }

    /// <summary>One resolved text style: font, size, minimum size, weight, line height and colour.</summary>
    internal sealed class PowerPointTextStyleSpec {
        internal PowerPointTextStyleSpec(PowerPointTextRole role, bool headingFont, double size, double minimum,
            bool bold, double lineHeight, string color, bool upper = false, double tracking = 0D) {
            Role = role;
            HeadingFont = headingFont;
            Size = size;
            Minimum = minimum;
            Bold = bold;
            LineHeight = lineHeight;
            Color = color;
            Upper = upper;
            Tracking = tracking;
        }

        internal PowerPointTextRole Role { get; }
        internal bool HeadingFont { get; }
        internal double Size { get; }
        internal double Minimum { get; }
        internal bool Bold { get; }
        internal double LineHeight { get; }
        internal string Color { get; }
        internal bool Upper { get; }

        /// <summary>Letter spacing in em.</summary>
        internal double Tracking { get; }
    }

    /// <summary>
    ///     Resolved design tokens for slides: palette roles, derived roles, fonts and the slide type scale.
    ///     Every slide colour, font and size the layout system writes comes from here, and every value here comes
    ///     from an <see cref="OfficeTheme" />.
    /// </summary>
    internal sealed class PowerPointLayoutTheme {
        private Dictionary<PowerPointTextRole, PowerPointTextStyleSpec> _styles = new();

        private PowerPointLayoutTheme(OfficeTheme theme) {
            Theme = theme;
            OfficeThemePalette p = theme.Palette;
            Background = Hex(p.Background);
            Text = Hex(p.Text);
            TextMuted = Hex(p.TextMuted);
            Accent1 = Hex(p.Accent1);
            Accent2 = Hex(p.Accent2);
            Accent3 = Hex(p.Accent3);
            Accent4 = Hex(p.Accent4);
            Accent5 = Hex(p.Accent5);
            Accent6 = Hex(p.Accent6);
            Success = Hex(p.Success);
            Warning = Hex(p.Warning);
            Danger = Hex(p.Danger);
            Surface = Hex(theme.TableStyle.BandFill.Resolve(theme));
            Hairline = Hex(theme.TableStyle.BorderColor.Resolve(theme));
            bool textDarker = Luminance(Text) < Luminance(Background);
            Ink = textDarker ? Text : Background;
            OnInk = textDarker ? Background : Text;
            OnAccent = Contrast(Accent1, "FFFFFF") >= 4.5D ? "FFFFFF" : Ink;
            HeadingFont = theme.Fonts.Heading.Family;
            BodyFont = theme.Fonts.Body.Family;
            TableHeaderFill = Hex(theme.TableStyle.HeaderFill.Resolve(theme));
            TableHeaderText = Hex(theme.TableStyle.HeaderText.Resolve(theme));
            TableBanding = theme.TableStyle.Banding;
            TableBorderWidth = theme.TableStyle.BorderWidth.ToPoints();
            TableCellPadding = theme.TableStyle.CellPadding.ToPoints();
            MarginLeft = theme.Spacing.SlideMarginLeft.ToPoints();
            MarginRight = theme.Spacing.SlideMarginRight.ToPoints();
            MarginTop = theme.Spacing.SlideMarginTop.ToPoints();
            ListIndent = theme.Spacing.ListIndent.ToPoints();
            ParagraphAfter = theme.Spacing.ParagraphAfter.ToPoints();

            OfficeThemeTypeScale s = theme.TypeScale;
            double title = s.Title.ToPoints(), body = s.Body.ToPoints(), caption = s.Caption.ToPoints();
            Add(PowerPointTextRole.Display, true, title * 1.9, 40, true, 1.0, OnAccent, tracking: -0.02);
            Add(PowerPointTextRole.SectionTitle, true, title * 1.5, 32, true, 1.04, OnInk, tracking: -0.015);
            Add(PowerPointTextRole.SectionNumeral, true, title * 8.0, 160, true, 0.8, OnInk, tracking: -0.04);
            Add(PowerPointTextRole.Title, true, title * 1.143, 26, true, 1.05, Text, tracking: -0.01);
            Add(PowerPointTextRole.Subtitle, false, body * 1.818, 16, false, 1.25, TextMuted);
            Add(PowerPointTextRole.Lede, false, body * 1.818, 16, false, 1.3, Text);
            Add(PowerPointTextRole.Eyebrow, false, caption * 1.333, 10, true, 1.2, Accent1, upper: true, tracking: 0.12);
            Add(PowerPointTextRole.Heading, true, body * 1.818, 16, true, 1.15, Text);
            Add(PowerPointTextRole.Body, false, body * 1.636, 16, false, 1.3, Text);
            Add(PowerPointTextRole.BodySmall, false, body * 1.364, 12, false, 1.35, TextMuted);
            Add(PowerPointTextRole.KpiValue, true, title * 1.714, 32, true, 1.0, Text, tracking: -0.02);
            Add(PowerPointTextRole.KpiLabel, false, body * 1.364, 12, false, 1.3, TextMuted);
            Add(PowerPointTextRole.Quote, true, title * 1.286, 24, false, 1.2, Text, tracking: -0.01);
            Add(PowerPointTextRole.TableHeader, false, body * 1.182, 11, true, 1.2, TableHeaderText);
            Add(PowerPointTextRole.TableBody, false, body * 1.273, 11, false, 1.2, Text);
            Add(PowerPointTextRole.Caption, false, caption * 1.222, 9, false, 1.3, TextMuted);
            Add(PowerPointTextRole.Footer, false, caption * 1.111, 9, false, 1.0, TextMuted);
        }

        internal OfficeTheme Theme { get; }
        internal string Background { get; }
        internal string Text { get; }
        internal string TextMuted { get; }
        internal string Accent1 { get; }
        internal string Accent2 { get; }
        internal string Accent3 { get; }
        internal string Accent4 { get; }
        internal string Accent5 { get; }
        internal string Accent6 { get; }
        internal string Success { get; }
        internal string Warning { get; }
        internal string Danger { get; }

        /// <summary>Tinted panel behind cards, quote slides and maps.</summary>
        internal string Surface { get; }

        /// <summary>Rules, tile borders and connectors.</summary>
        internal string Hairline { get; }

        /// <summary>The darker of text and background.</summary>
        internal string Ink { get; }

        /// <summary>The lighter of text and background.</summary>
        internal string OnInk { get; }

        /// <summary>Text on an accent1 background.</summary>
        internal string OnAccent { get; }

        internal string HeadingFont { get; private set; }
        internal string BodyFont { get; private set; }
        internal string TableHeaderFill { get; }
        internal string TableHeaderText { get; }
        internal bool TableBanding { get; }
        internal double TableBorderWidth { get; }
        internal double TableCellPadding { get; private set; }
        internal double MarginLeft { get; private set; }
        internal double MarginRight { get; private set; }
        internal double MarginTop { get; private set; }
        internal double ListIndent { get; private set; }
        internal double ParagraphAfter { get; private set; }

        internal PowerPointTextStyleSpec this[PowerPointTextRole role] => _styles[role];

        internal string Accent(int index) => index switch {
            2 => Accent2,
            3 => Accent3,
            4 => Accent4,
            5 => Accent5,
            6 => Accent6,
            _ => Accent1
        };

        internal static PowerPointLayoutTheme From(OfficeTheme? theme) => new(theme ?? OfficeTheme.Default);

        /// <summary>The same tokens measured with other font families (the fonts a template's theme part names).</summary>
        internal PowerPointLayoutTheme WithFontFamilies(string? heading, string? body) {
            if ((heading == null || heading == HeadingFont) && (body == null || body == BodyFont)) return this;
            var copy = (PowerPointLayoutTheme)MemberwiseClone();
            copy.HeadingFont = heading ?? HeadingFont;
            copy.BodyFont = body ?? BodyFont;
            return copy;
        }

        private const double ReadableFloor = 10D;

        /// <summary>Size of the slide relative to the reference the type scale is designed for.</summary>
        internal double Scale { get; private set; } = 1D;

        /// <summary>
        ///     The same tokens for a smaller slide: text sizes, minimums, margins and spacing scale by
        ///     <paramref name="factor" />, with a 10 pt floor (the preflight's readable minimum) so text stays readable.
        /// </summary>
        internal PowerPointLayoutTheme Scaled(double factor) {
            if (factor >= 0.97D || Math.Abs(factor - Scale) < 0.001D) return this;
            var copy = (PowerPointLayoutTheme)MemberwiseClone();
            double relative = factor / Scale;
            copy._styles = new Dictionary<PowerPointTextRole, PowerPointTextStyleSpec>();
            foreach (KeyValuePair<PowerPointTextRole, PowerPointTextStyleSpec> pair in _styles) {
                PowerPointTextStyleSpec style = pair.Value;
                double floor = Math.Min(style.Size, ReadableFloor);
                double size = Math.Max(floor, RoundSize(style.Size * relative));
                double minimum = Math.Min(size, Math.Max(floor, Math.Round(style.Minimum * relative)));
                copy._styles[pair.Key] = new PowerPointTextStyleSpec(style.Role, style.HeadingFont, size, minimum, style.Bold,
                    style.LineHeight, style.Color, style.Upper, style.Tracking);
            }

            copy.MarginLeft = MarginLeft * relative;
            copy.MarginRight = MarginRight * relative;
            copy.MarginTop = MarginTop * relative;
            copy.ListIndent = ListIndent * relative;
            copy.ParagraphAfter = ParagraphAfter * relative;
            copy.TableCellPadding = TableCellPadding * relative;
            copy.Scale = factor;
            return copy;
        }

        /// <summary>
        ///     Maps a designer theme onto the default theme: the brand accents and fonts replace the theme's, the
        ///     type scale, spacing and neutral roles stay. Colours that would fail contrast are not taken over.
        /// </summary>
        internal static PowerPointLayoutTheme From(PowerPointDesignTheme design, OfficeTheme? baseTheme = null) {
            if (design == null) throw new ArgumentNullException(nameof(design));
            OfficeTheme theme = baseTheme ?? OfficeTheme.Default;
            OfficeThemePalette palette = theme.Palette;
            OfficeColor accent = Color(design.AccentColor, palette.Accent1);
            // An accent that cannot carry text on the background (e.g. a pastel brand colour) is darkened until it can.
            accent = EnsureContrast(accent, palette.Background, 4.5D);
            OfficeColor accent2 = EnsureContrast(Color(design.Accent2Color, palette.Accent2), palette.Background, 3D);
            OfficeColor accent3 = EnsureContrast(Color(design.Accent3Color, palette.Accent3), palette.Background, 3D);
            // Series and comparison columns are told apart by these accents. A brand accent that is only a shade of
            // the primary keeps the theme's own accent instead, so two chart series never share one blue.
            // A design that repeats the theme's own primary accent composes with the theme as it is.
            if (accent.ToRgbHex() == palette.Accent1.ToRgbHex()) {
                accent2 = palette.Accent2;
                accent3 = palette.Accent3;
            }
            if (!Distinct(accent2, accent)) accent2 = Distinct(palette.Accent2, accent) ? palette.Accent2 : palette.Accent4;
            if (!Distinct(accent3, accent) || !Distinct(accent3, accent2)) accent3 = palette.Accent3;
            palette = palette.With(accent1: accent, accent2: accent2, accent3: accent3);
            OfficeThemeFont heading = theme.Fonts.Heading.With(family: Font(design.HeadingFontName, theme.Fonts.Heading.Family));
            OfficeThemeFont body = theme.Fonts.Body.With(family: Font(design.BodyFontName, theme.Fonts.Body.Family));
            OfficeTheme mapped = theme.WithPalette(palette).WithFonts(new OfficeThemeFonts(heading, body));
            return new PowerPointLayoutTheme(mapped);
        }

        /// <summary>Whether two accents read as different series: a clear hue step or a clear lightness step.</summary>
        internal static bool Distinct(OfficeColor a, OfficeColor b) {
            static (double H, double L) HueLightness(OfficeColor c) {
                double r = c.R / 255D, g = c.G / 255D, bl = c.B / 255D;
                double max = Math.Max(r, Math.Max(g, bl)), min = Math.Min(r, Math.Min(g, bl)), d = max - min;
                double h = d == 0D ? -1D : max == r ? 60D * (((g - bl) / d + 6D) % 6D) : max == g ? 60D * ((bl - r) / d + 2D) : 60D * ((r - g) / d + 4D);
                return (h, (max + min) / 2D);
            }
            (double ha, double la) = HueLightness(a);
            (double hb, double lb) = HueLightness(b);
            double hue = ha < 0D || hb < 0D ? 180D : Math.Min(Math.Abs(ha - hb), 360D - Math.Abs(ha - hb));
            return hue >= 40D || Math.Abs(la - lb) >= 0.25D;
        }

        private void Add(PowerPointTextRole role, bool headingFont, double size, double minimum, bool bold,
            double lineHeight, string color, bool upper = false, double tracking = 0D) {
            _styles[role] = new PowerPointTextStyleSpec(role, headingFont, RoundSize(size), minimum, bold, lineHeight,
                color, upper, tracking);
        }

        /// <summary>Sizes of 18 pt and above round to even points, smaller sizes to whole points.</summary>
        internal static double RoundSize(double value) =>
            value >= 18D ? Math.Round(value / 2D, MidpointRounding.AwayFromZero) * 2D : Math.Round(value, MidpointRounding.AwayFromZero);

        internal static string Hex(OfficeColor color) => color.ToRgbHex();

        private static OfficeColor Color(string? hex, OfficeColor fallback) {
            if (string.IsNullOrWhiteSpace(hex)) return fallback;
            try {
                return OfficeColor.ParseHex(hex!.Trim().TrimStart('#'));
            } catch (FormatException) {
                return fallback;
            } catch (ArgumentException) {
                return fallback;
            }
        }

        // Fonts that ship with Microsoft Office on Windows and macOS. Others (Poppins, Lato, Segoe UI on a Mac, the
        // cloud-only Aptos Display) are replaced by the theme font so a deck never depends on a font the reader lacks.
        private static readonly HashSet<string> PortableFonts = new(StringComparer.OrdinalIgnoreCase) {
            "Aptos", "Calibri", "Cambria", "Georgia", "Arial", "Times New Roman", "Verdana", "Tahoma", "Trebuchet MS", "Garamond", "Consolas"
        };

        private static string Font(string? name, string fallback) =>
            string.IsNullOrWhiteSpace(name) || !PortableFonts.Contains(name!.Trim()) ? fallback : name!.Trim();

        internal static OfficeColor EnsureContrast(OfficeColor color, OfficeColor background, double ratio) {
            string bg = Hex(background);
            OfficeColor current = color;
            for (int i = 0; i < 20 && Contrast(Hex(current), bg) < ratio; i++) {
                bool darken = Luminance(bg) > 0.5D;
                current = Mix(current, darken ? OfficeColor.ParseHex("000000") : OfficeColor.ParseHex("FFFFFF"), 0.08D);
            }

            return current;
        }

        internal static OfficeColor Mix(OfficeColor a, OfficeColor b, double amount) {
            byte M(byte x, byte y) => (byte)Math.Round(x + (y - x) * amount);
            return OfficeColor.ParseHex(string.Format(CultureInfo.InvariantCulture, "{0:X2}{1:X2}{2:X2}",
                M(a.R, b.R), M(a.G, b.G), M(a.B, b.B)));
        }

        /// <summary>Mixes <paramref name="hex" /> over <paramref name="over" /> at <paramref name="alpha" />.</summary>
        internal static string Blend(string hex, string over, double alpha) =>
            Hex(Mix(OfficeColor.ParseHex(over), OfficeColor.ParseHex(hex), alpha));

        /// <summary>
        ///     <paramref name="hex" /> blended over <paramref name="background" /> at <paramref name="alpha" />, made more
        ///     opaque when needed so the result still reaches <paramref name="ratio" /> (WCAG AA for text by default).
        /// </summary>
        internal static string Muted(string hex, string background, double alpha, double ratio = 4.5D) {
            double a = alpha;
            string result = Blend(hex, background, a);
            while (a < 1D && Contrast(result, background) < ratio) {
                a = Math.Min(1D, a + 0.02D);
                result = Blend(hex, background, a);
            }

            return result;
        }

        internal static double Luminance(string hex) {
            OfficeColor c = OfficeColor.ParseHex(hex);
            static double F(byte v) {
                double s = v / 255D;
                return s <= 0.03928D ? s / 12.92D : Math.Pow((s + 0.055D) / 1.055D, 2.4D);
            }

            return 0.2126D * F(c.R) + 0.7152D * F(c.G) + 0.0722D * F(c.B);
        }

        internal static double Contrast(string a, string b) {
            double la = Luminance(a), lb = Luminance(b);
            if (la < lb) (la, lb) = (lb, la);
            return (la + 0.05D) / (lb + 0.05D);
        }
    }
}
