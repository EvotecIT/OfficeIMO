using System.Globalization;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     PANOSE and pitch-family classification for theme and run fonts. Renderers that do not have the named
    ///     font pick a substitute from this classification, so a stale Office panose (or none at all) makes a
    ///     missing sans-serif font fall back to a serif or decorative face.
    /// </summary>
    internal static class PowerPointFontClassification {
        // pitchFamily: low nibble 2 = variable pitch; high nibble family: 0x10 roman (serif), 0x20 swiss (sans), 0x30 modern (mono).
        private const byte SansSerif = 0x22;
        private const byte Serif = 0x12;
        private const byte Monospace = 0x31;

        private static readonly Dictionary<string, (string Panose, byte PitchFamily)> Known =
            new(StringComparer.OrdinalIgnoreCase) {
                ["Aptos"] = ("020B0004020202020204", SansSerif),
                ["Aptos Display"] = ("020B0004020202020204", SansSerif),
                ["Calibri"] = ("020F0502020204030204", SansSerif),
                ["Calibri Light"] = ("020F0302020204030204", SansSerif),
                ["Carlito"] = ("020F0502020204030204", SansSerif),
                ["Arial"] = ("020B0604020202020204", SansSerif),
                ["Liberation Sans"] = ("020B0604020202020204", SansSerif),
                ["Helvetica"] = ("020B0604020202020204", SansSerif),
                ["Helvetica Neue"] = ("02000503000000020004", SansSerif),
                ["Segoe UI"] = ("020B0502040204020203", SansSerif),
                ["Lato"] = ("020F0502020204030203", SansSerif),
                ["Poppins"] = ("00000500000000000000", SansSerif),
                ["Inter"] = ("020B0502030000000004", SansSerif),
                ["Verdana"] = ("020B0604030504040204", SansSerif),
                ["Tahoma"] = ("020B0604030504040204", SansSerif),
                ["Georgia"] = ("02040502050405020303", Serif),
                ["Gelasio"] = ("02040502050405020303", Serif),
                ["Cambria"] = ("02040503050406030204", Serif),
                ["Caladea"] = ("02040503050406030204", Serif),
                ["Times New Roman"] = ("02020603050405020304", Serif),
                ["Liberation Serif"] = ("02020603050405020304", Serif),
                ["Garamond"] = ("02020404030301010803", Serif),
                ["Consolas"] = ("020B0609020204030204", Monospace),
                ["Courier New"] = ("02070309020205020404", Monospace),
            };

        /// <summary>Writes the classification that matches <paramref name="typeface" />, removing a stale one.</summary>
        /// <param name="font">A text font element (<c>a:latin</c>, <c>a:ea</c>, <c>a:cs</c> or <c>a:sym</c>).</param>
        /// <param name="typeface">Typeface to write.</param>
        internal static void Apply(OfficeXmlNode font, string typeface) {
            if (font.IsNone) throw new ArgumentNullException(nameof(font));
            font.SetAttribute("typeface", typeface);
            Classify(typeface, out string? panose, out sbyte? pitchFamily, out bool characterSet);
            font.SetAttribute("panose", panose);
            font.SetAttribute("pitchFamily", pitchFamily?.ToString(CultureInfo.InvariantCulture));
            font.SetAttribute("charset", characterSet ? "0" : null);
        }

        /// <summary>The same as <see cref="Apply(OfficeXmlNode, string)"/> for a font element of an SDK part, passed as <see cref="object"/>.</summary>
        internal static void Apply(object font, string typeface) {
            ArgumentNullException.ThrowIfNull(font);
            Apply(OfficeXmlNode.FromSdk(font), typeface);
        }

        /// <summary>
        ///     The classification <see cref="Apply(OfficeXmlNode, string)"/> writes: the PANOSE (null when it is
        ///     unknown), the pitch family and whether the character set is written. A theme font reference
        ///     (<c>+mj-lt</c>) carries none of them.
        /// </summary>
        internal static void Classify(string typeface, out string? panose, out sbyte? pitchFamily, out bool characterSet) {
            if (typeface.StartsWith("+", StringComparison.Ordinal)) {
                panose = null;
                pitchFamily = null;
                characterSet = false;
                return;
            }

            characterSet = true;
            if (Known.TryGetValue(typeface, out (string Panose, byte PitchFamily) value)) {
                panose = value.Panose;
                pitchFamily = unchecked((sbyte)value.PitchFamily);
                return;
            }

            // Unknown family: a guessed classification is worse than none, but keep the generic family hint.
            panose = null;
            pitchFamily = IsLikelySerif(typeface) ? unchecked((sbyte)Serif) : unchecked((sbyte)SansSerif);
        }

        internal static bool IsLikelySerif(string typeface) {
            if (Known.TryGetValue(typeface, out (string Panose, byte PitchFamily) value)) return value.PitchFamily == Serif;
            return typeface.IndexOf("Serif", StringComparison.OrdinalIgnoreCase) >= 0 && typeface.IndexOf("Sans", StringComparison.OrdinalIgnoreCase) < 0
                || typeface.IndexOf("Times", StringComparison.OrdinalIgnoreCase) >= 0
                || typeface.IndexOf("Garamond", StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
