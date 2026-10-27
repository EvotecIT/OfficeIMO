using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private PowerPointLayoutTheme? _layoutTheme;
        private OfficeTheme? _baseTheme;

        /// <summary>The theme a new presentation was created with; the default theme for loaded presentations.</summary>
        internal OfficeTheme BaseTheme => _baseTheme ?? OfficeTheme.Default;

        /// <summary>True when OfficeIMO wrote this presentation's master and layouts, so it may rewrite them.</summary>
        internal bool OwnsLayoutTheme => _baseTheme != null;
        private PowerPointLayoutDeck? _layoutDeck;

        /// <summary>Writes the theme part, master text styles, placeholders and layouts of a new presentation.</summary>
        internal void ApplyLayoutTheme(OfficeTheme theme) {
            _baseTheme = theme;
            _layoutTheme = PowerPointLayoutTheme.From(theme);
            PowerPointLayoutMasterWriter.Apply(_presentationPart, _layoutTheme);
        }

        /// <summary>Rewrites the theme-derived parts for a designer theme mapped onto the presentation's theme.</summary>
        internal void ApplyLayoutTheme(PowerPointLayoutTheme theme) {
            _layoutTheme = theme;
            PowerPointLayoutMasterWriter.Apply(_presentationPart, theme);
        }

        /// <summary>
        ///     Tokens used by the slide layout system. New presentations use the theme they were created with; a
        ///     loaded presentation keeps its own colours and fonts on top of the default type scale and grid.
        /// </summary>
        internal PowerPointLayoutTheme LayoutTheme => _layoutTheme ??= ReadLayoutTheme();

        /// <summary>Deck-level state of the layout system: section titles, numbering and footer text.</summary>
        internal PowerPointLayoutDeck LayoutDeck => _layoutDeck ??= new PowerPointLayoutDeck(this);

        private PowerPointLayoutTheme ReadLayoutTheme() {
            OfficeTheme theme = OfficeTheme.Default;
            SlideMasterPart? master = _presentationPart.SlideMasterParts.FirstOrDefault();
            A.ThemeElements? elements = master?.ThemePart?.Theme?.ThemeElements;
            if (elements == null) return PowerPointLayoutTheme.From(theme);

            OfficeColor? Read(DocumentFormat.OpenXml.OpenXmlCompositeElement? slot) {
                string? hex = slot?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value
                              ?? slot?.GetFirstChild<A.SystemColor>()?.LastColor?.Value;
                if (string.IsNullOrWhiteSpace(hex) || hex!.Length != 6) return null;
                try {
                    return OfficeColor.ParseHex(hex);
                } catch (FormatException) {
                    return null;
                }
            }

            A.ColorScheme? colors = elements.ColorScheme;
            OfficeThemePalette palette = theme.Palette;
            OfficeColor? text = Read(colors?.Dark1Color);
            OfficeColor? background = Read(colors?.Light1Color);
            palette = palette.With(
                background: background,
                text: text,
                accent1: Read(colors?.Accent1Color),
                accent2: Read(colors?.Accent2Color),
                accent3: Read(colors?.Accent3Color),
                accent4: Read(colors?.Accent4Color),
                accent5: Read(colors?.Accent5Color),
                accent6: Read(colors?.Accent6Color));
            string? heading = elements.FontScheme?.MajorFont?.LatinFont?.Typeface?.Value;
            string? body = elements.FontScheme?.MinorFont?.LatinFont?.Typeface?.Value;
            OfficeThemeFonts fonts = new(
                string.IsNullOrWhiteSpace(heading) ? theme.Fonts.Heading : theme.Fonts.Heading.With(family: heading),
                string.IsNullOrWhiteSpace(body) ? theme.Fonts.Body : theme.Fonts.Body.With(family: body));
            return PowerPointLayoutTheme.From(theme.WithPalette(palette).WithFonts(fonts));
        }
    }
}
