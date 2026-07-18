using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents the eight-color scheme stored by PowerPoint 97-2003.</summary>
    public sealed class LegacyPptColorScheme {
        internal LegacyPptColorScheme(IReadOnlyList<string> colors) {
            if (colors == null) throw new ArgumentNullException(nameof(colors));
            if (colors.Count != 8) throw new ArgumentException("A legacy PowerPoint color scheme contains eight colors.", nameof(colors));
            Background = colors[0];
            Text = colors[1];
            Shadow = colors[2];
            TitleText = colors[3];
            Fill = colors[4];
            Accent1 = colors[5];
            Accent2 = colors[6];
            Accent3 = colors[7];
        }

        /// <summary>Gets the background color as RRGGBB.</summary>
        public string Background { get; }
        /// <summary>Gets the text color as RRGGBB.</summary>
        public string Text { get; }
        /// <summary>Gets the shadow color as RRGGBB.</summary>
        public string Shadow { get; }
        /// <summary>Gets the title-text color as RRGGBB.</summary>
        public string TitleText { get; }
        /// <summary>Gets the default fill color as RRGGBB.</summary>
        public string Fill { get; }
        /// <summary>Gets accent color 1 as RRGGBB.</summary>
        public string Accent1 { get; }
        /// <summary>Gets accent color 2 as RRGGBB.</summary>
        public string Accent2 { get; }
        /// <summary>Gets accent color 3 as RRGGBB.</summary>
        public string Accent3 { get; }

        /// <summary>Tries to get a color by its zero-based binary PowerPoint scheme index.</summary>
        public bool TryGetColor(byte index, out string? color) {
            color = index switch {
                0 => Background,
                1 => Text,
                2 => Shadow,
                3 => TitleText,
                4 => Fill,
                5 => Accent1,
                6 => Accent2,
                7 => Accent3,
                _ => null
            };
            return color != null;
        }

        internal OfficeColor? ResolveOfficeArtColor(byte index) =>
            TryGetColor(index, out string? color) && OfficeColor.TryParseHex(color, out OfficeColor parsed)
                ? parsed
                : null;
    }
}
