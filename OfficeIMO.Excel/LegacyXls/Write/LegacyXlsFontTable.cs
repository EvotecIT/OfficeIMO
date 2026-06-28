using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal sealed class LegacyXlsFontTable {
        private const ushort AutomaticColorIndex = 0x7fff;
        private const string DefaultFontName = "Calibri";
        private const double DefaultFontSize = 11d;

        private readonly Dictionary<FontKey, ushort> _indexes = new();
        private readonly List<FontKey> _entries = new();
        private readonly Dictionary<string, ushort> _customColorIndexes = new(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<int> _reservedPaletteSlots = new();
        private readonly LegacyXlsThemeColorResolver _themeColors;
        private List<string>? _paletteColors;

        private LegacyXlsFontTable(LegacyXlsThemeColorResolver themeColors) {
            _themeColors = themeColors ?? throw new ArgumentNullException(nameof(themeColors));
            AddFont(new FontKey(DefaultFontName, ToTwips(DefaultFontSize), AutomaticColorIndex, bold: false, italic: false, underlineStyle: 0, strikeout: false, outline: false, shadow: false, condense: false, extend: false, LegacyXlsFontEscapement.None, family: 0, characterSet: 1));
        }

        internal IReadOnlyList<byte[]> FontRecords => _entries.Select(BuildFontPayload).ToArray();

        internal byte[]? PaletteRecord => _paletteColors == null ? null : BuildPalettePayload(_paletteColors);

        internal static LegacyXlsFontTable Create(ExcelDocument document) {
            return new LegacyXlsFontTable(LegacyXlsThemeColorResolver.Create(document));
        }

        internal bool TryGetFontIndex(RunProperties? properties, out ushort fontIndex, out string? reason) {
            reason = null;
            fontIndex = 0;
            if (properties == null) {
                return true;
            }

            if (!properties.HasChildren && !properties.GetAttributes().Any()) {
                return true;
            }

            if (!SupportsRunFontPropertyMetadata(properties, out reason)
                || !SupportsSingleFontProperty<RunFont>(properties, "rich-text", "font name", out reason)
                || !SupportsSingleFontProperty<FontSize>(properties, "rich-text", "font size", out reason)
                || !SupportsSingleFontProperty<DocumentFormat.OpenXml.Spreadsheet.Color>(properties, "rich-text", "font color", out reason)
                || !SupportsSingleFontProperty<Bold>(properties, "rich-text", "bold", out reason)
                || !SupportsSingleFontProperty<Italic>(properties, "rich-text", "italic", out reason)
                || !SupportsSingleFontProperty<Underline>(properties, "rich-text", "underline", out reason)
                || !SupportsSingleFontProperty<Strike>(properties, "rich-text", "strike", out reason)
                || !SupportsSingleFontProperty<Outline>(properties, "rich-text", "outline", out reason)
                || !SupportsSingleFontProperty<Shadow>(properties, "rich-text", "shadow", out reason)
                || !SupportsSingleFontProperty<Condense>(properties, "rich-text", "condense", out reason)
                || !SupportsSingleFontProperty<Extend>(properties, "rich-text", "extend", out reason)
                || !SupportsSingleFontProperty<VerticalTextAlignment>(properties, "rich-text", "vertical text alignment", out reason)
                || !SupportsSingleFontProperty(properties, "rich-text", "family", "font family", out reason)
                || !SupportsSingleFontProperty(properties, "rich-text", "charset", "font character set", out reason)) {
                return false;
            }

            string fontName = properties.GetFirstChild<RunFont>()?.Val?.Value ?? DefaultFontName;
            if (string.IsNullOrWhiteSpace(fontName)) {
                fontName = DefaultFontName;
            }

            if (fontName.Length > byte.MaxValue) {
                reason = "rich-text comment font names longer than 255 characters";
                return false;
            }

            double fontSize = properties.GetFirstChild<FontSize>()?.Val?.Value ?? DefaultFontSize;
            if (fontSize <= 0 || fontSize * 20d > ushort.MaxValue) {
                reason = "rich-text comment font sizes outside BIFF8 limits";
                return false;
            }

            if (!TryGetColorIndex(properties.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>(), "rich-text comment", "font colors", out ushort colorIndex, out reason)) {
                return false;
            }

            if (!TryGetFontByteProperty(properties, "family", "rich-text", "font family", 0, out byte family, out reason)
                || !TryGetFontByteProperty(properties, "charset", "rich-text", "font character set", 1, out byte characterSet, out reason)
                || !TryGetUnderlineStyle(properties.GetFirstChild<Underline>(), "rich-text", out byte underlineStyle, out reason)
                || !TryGetFontEscapement(properties.GetFirstChild<VerticalTextAlignment>(), "rich-text", out LegacyXlsFontEscapement escapement, out reason)) {
                return false;
            }

            var key = new FontKey(
                fontName.Trim(),
                ToTwips(fontSize),
                colorIndex,
                IsEnabled(properties.GetFirstChild<Bold>()),
                IsEnabled(properties.GetFirstChild<Italic>()),
                underlineStyle,
                IsEnabled(properties.GetFirstChild<Strike>()),
                IsEnabled(properties.GetFirstChild<Outline>()),
                IsEnabled(properties.GetFirstChild<Shadow>()),
                IsEnabled(properties.GetFirstChild<Condense>()),
                IsEnabled(properties.GetFirstChild<Extend>()),
                escapement,
                family,
                characterSet);

            fontIndex = GetOrAddFont(key);
            return true;
        }

        internal bool TryGetFontIndex(Font? font, out ushort fontIndex, out string? reason) {
            reason = null;
            fontIndex = 0;
            if (font == null) {
                return true;
            }

            if (!font.HasChildren && !font.GetAttributes().Any()) {
                return true;
            }

            if (!SupportsCellFontPropertyMetadata(font, out reason)
                || !SupportsSingleFontProperty<FontName>(font, "cell", "font name", out reason)
                || !SupportsSingleFontProperty<FontSize>(font, "cell", "font size", out reason)
                || !SupportsSingleFontProperty<DocumentFormat.OpenXml.Spreadsheet.Color>(font, "cell", "font color", out reason)
                || !SupportsSingleFontProperty<Bold>(font, "cell", "bold", out reason)
                || !SupportsSingleFontProperty<Italic>(font, "cell", "italic", out reason)
                || !SupportsSingleFontProperty<Underline>(font, "cell", "underline", out reason)
                || !SupportsSingleFontProperty<Strike>(font, "cell", "strike", out reason)
                || !SupportsSingleFontProperty<Outline>(font, "cell", "outline", out reason)
                || !SupportsSingleFontProperty<Shadow>(font, "cell", "shadow", out reason)
                || !SupportsSingleFontProperty<Condense>(font, "cell", "condense", out reason)
                || !SupportsSingleFontProperty<Extend>(font, "cell", "extend", out reason)
                || !SupportsSingleFontProperty<VerticalTextAlignment>(font, "cell", "vertical text alignment", out reason)
                || !SupportsSingleFontProperty(font, "cell", "family", "font family", out reason)
                || !SupportsSingleFontProperty(font, "cell", "charset", "font character set", out reason)) {
                return false;
            }

            string fontName = font.GetFirstChild<FontName>()?.Val?.Value ?? DefaultFontName;
            if (string.IsNullOrWhiteSpace(fontName)) {
                fontName = DefaultFontName;
            }

            if (fontName.Length > byte.MaxValue) {
                reason = "cell font names longer than 255 characters";
                return false;
            }

            double fontSize = font.GetFirstChild<FontSize>()?.Val?.Value ?? DefaultFontSize;
            if (fontSize <= 0 || fontSize * 20d > ushort.MaxValue) {
                reason = "cell font sizes outside BIFF8 limits";
                return false;
            }

            if (!TryGetColorIndex(font.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>(), "cell", "font colors", out ushort colorIndex, out reason)) {
                return false;
            }

            if (!TryGetFontByteProperty(font, "family", "cell", "font family", 0, out byte family, out reason)
                || !TryGetFontByteProperty(font, "charset", "cell", "font character set", 1, out byte characterSet, out reason)
                || !TryGetUnderlineStyle(font.GetFirstChild<Underline>(), "cell", out byte underlineStyle, out reason)
                || !TryGetFontEscapement(font.GetFirstChild<VerticalTextAlignment>(), "cell", out LegacyXlsFontEscapement escapement, out reason)) {
                return false;
            }

            var key = new FontKey(
                fontName.Trim(),
                ToTwips(fontSize),
                colorIndex,
                IsEnabled(font.GetFirstChild<Bold>()),
                IsEnabled(font.GetFirstChild<Italic>()),
                underlineStyle,
                IsEnabled(font.GetFirstChild<Strike>()),
                IsEnabled(font.GetFirstChild<Outline>()),
                IsEnabled(font.GetFirstChild<Shadow>()),
                IsEnabled(font.GetFirstChild<Condense>()),
                IsEnabled(font.GetFirstChild<Extend>()),
                escapement,
                family,
                characterSet);

            fontIndex = GetOrAddFont(key);
            return true;
        }

        private ushort GetOrAddFont(FontKey key) {
            return _indexes.TryGetValue(key, out ushort existingIndex)
                ? existingIndex
                : AddFont(key);
        }

        private ushort AddFont(FontKey key) {
            int recordIndex = _entries.Count;
            ushort externalIndex = checked((ushort)(recordIndex < 4 ? recordIndex : recordIndex + 1));
            _entries.Add(key);
            _indexes.Add(key, externalIndex);
            return externalIndex;
        }

        internal bool TryGetColorIndex(ColorType? color, string context, string subject, out ushort colorIndex, out string? reason) {
            colorIndex = AutomaticColorIndex;
            reason = null;
            if (color == null || color.Auto?.Value == true) {
                return true;
            }

            if (color.Indexed?.Value is uint indexedColor) {
                if (indexedColor > ushort.MaxValue) {
                    reason = $"{context} indexed {subject} outside BIFF8 limits";
                    return false;
                }

                colorIndex = checked((ushort)indexedColor);
                return true;
            }

            if (!string.IsNullOrWhiteSpace(color.Rgb?.Value)) {
                string? argb = NormalizeArgb(color.Rgb!.Value!);
                if (argb != null && color.Tint?.Value is double tint) {
                    argb = ApplyTint(argb, tint);
                }

                return TryMapArgbColor(argb, out colorIndex, out reason);
            }

            if (color.Theme?.Value is uint themeIndex) {
                if (!_themeColors.TryResolve(themeIndex, color.Tint?.Value, out string? argb)) {
                    reason = $"unknown theme {context} {subject}";
                    return false;
                }

                return TryMapArgbColor(argb, out colorIndex, out reason);
            }

            if (color.Tint != null) {
                reason = $"tint-only {context} {subject}";
                return false;
            }

            return true;
        }

        private bool TryMapArgbColor(string? argb, out ushort colorIndex, out string? reason) {
            colorIndex = AutomaticColorIndex;
            reason = null;
            if (argb == null) {
                reason = "invalid rich-text comment font colors";
                return false;
            }

            if (BiffColorPalette.TryGetBuiltInColorIndex(argb, out colorIndex)) {
                return true;
            }

            if (BiffColorPalette.TryGetDefaultPaletteColorIndex(argb, out colorIndex)) {
                int paletteSlot = colorIndex - 8;
                if (_paletteColors == null
                    || !_reservedPaletteSlots.Contains(paletteSlot)
                    || string.Equals(_paletteColors[paletteSlot], argb, StringComparison.OrdinalIgnoreCase)) {
                    _reservedPaletteSlots.Add(paletteSlot);
                    if (_paletteColors != null) {
                        _paletteColors[paletteSlot] = argb;
                    }

                    return true;
                }

                return TryAssignPaletteColor(argb, out colorIndex, out reason);
            }

            return TryAssignPaletteColor(argb, out colorIndex, out reason);
        }

        private bool TryAssignPaletteColor(string argb, out ushort colorIndex, out string? reason) {
            reason = null;
            if (_customColorIndexes.TryGetValue(argb, out colorIndex)) {
                return true;
            }

            _paletteColors ??= BiffColorPalette.DefaultPaletteColors.ToList();
            for (int slot = 0; slot < _paletteColors.Count; slot++) {
                if (_reservedPaletteSlots.Contains(slot)) {
                    continue;
                }

                _reservedPaletteSlots.Add(slot);
                _paletteColors[slot] = argb;
                colorIndex = checked((ushort)(8 + slot));
                _customColorIndexes.Add(argb, colorIndex);
                return true;
            }

            reason = "more than 56 distinct non-built-in rich-text comment font colors";
            colorIndex = AutomaticColorIndex;
            return false;
        }

        private static string? NormalizeArgb(string value) {
            string hex = value.Trim().TrimStart('#');
            if (hex.Length == 6) {
                hex = "FF" + hex;
            }

            if (hex.Length != 8 || hex.Any(ch => !Uri.IsHexDigit(ch))) {
                return null;
            }

            return hex.ToUpperInvariant();
        }

        private static string ApplyTint(string argb, double tint) {
            byte alpha = Convert.ToByte(argb.Substring(0, 2), 16);
            byte red = Convert.ToByte(argb.Substring(2, 2), 16);
            byte green = Convert.ToByte(argb.Substring(4, 2), 16);
            byte blue = Convert.ToByte(argb.Substring(6, 2), 16);
            return alpha.ToString("X2", CultureInfo.InvariantCulture)
                + ApplyTintChannel(red, tint).ToString("X2", CultureInfo.InvariantCulture)
                + ApplyTintChannel(green, tint).ToString("X2", CultureInfo.InvariantCulture)
                + ApplyTintChannel(blue, tint).ToString("X2", CultureInfo.InvariantCulture);
        }

        private static byte ApplyTintChannel(byte channel, double tint) {
            double value = tint < 0D
                ? channel * (1D + tint)
                : channel * (1D - tint) + (255D * tint);
            return (byte)Math.Max(0D, Math.Min(255D, Math.Round(value)));
        }

        private static int ToTwips(double points) {
            return checked((int)Math.Round(points * 20d, MidpointRounding.AwayFromZero));
        }

        private static bool IsEnabled(BooleanPropertyType? element) {
            return element != null && element.Val?.Value != false;
        }

        private static bool SupportsSingleFontProperty<TElement>(OpenXmlElement properties, string context, string elementName, out string? reason)
            where TElement : OpenXmlElement {
            reason = null;
            if (properties.Elements<TElement>().Take(2).Count() <= 1) {
                return true;
            }

            reason = $"{context} font properties with duplicate {elementName} elements";
            return false;
        }

        private static bool SupportsSingleFontProperty(OpenXmlElement properties, string context, string elementName, string description, out string? reason) {
            reason = null;
            if (properties.ChildElements.Where(child => IsSpreadsheetElement(child, elementName)).Take(2).Count() <= 1) {
                return true;
            }

            reason = $"{context} font properties with duplicate {description} elements";
            return false;
        }

        private static bool SupportsRunFontPropertyMetadata(RunProperties properties, out string? reason) {
            reason = null;
            if (properties.GetAttributes().Any()) {
                reason = "rich-text font properties with unsupported metadata";
                return false;
            }

            if (properties.ChildElements.Any(child =>
                child is not RunFont
                && child is not FontSize
                && child is not DocumentFormat.OpenXml.Spreadsheet.Color
                && child is not Bold
                && child is not Italic
                && child is not Underline
                && child is not Strike
                && child is not Outline
                && child is not Shadow
                && child is not Condense
                && child is not Extend
                && child is not VerticalTextAlignment
                && !IsSpreadsheetElement(child, "family")
                && !IsSpreadsheetElement(child, "charset"))) {
                reason = "rich-text font properties with unsupported metadata";
                return false;
            }

            return true;
        }

        private static bool SupportsCellFontPropertyMetadata(Font font, out string? reason) {
            reason = null;
            if (font.GetAttributes().Any()) {
                reason = "cell font properties with unsupported metadata";
                return false;
            }

            if (font.ChildElements.Any(child =>
                child is not FontName
                && child is not FontSize
                && child is not DocumentFormat.OpenXml.Spreadsheet.Color
                && child is not Bold
                && child is not Italic
                && child is not Underline
                && child is not Strike
                && child is not Outline
                && child is not Shadow
                && child is not Condense
                && child is not Extend
                && child is not VerticalTextAlignment
                && !IsSpreadsheetElement(child, "family")
                && !IsSpreadsheetElement(child, "charset"))) {
                reason = "cell font properties with unsupported metadata";
                return false;
            }

            return true;
        }

        private static byte[] BuildFontPayload(FontKey key) {
            byte[] nameBytes = EncodeUnicodeString(key.Name, out byte flags);
            using var stream = new MemoryStream();
            WriteUInt16(stream, checked((ushort)key.SizeTwips));
            WriteUInt16(stream, BuildOptions(key));
            WriteUInt16(stream, key.ColorIndex);
            WriteUInt16(stream, key.Bold ? (ushort)700 : (ushort)400);
            WriteUInt16(stream, ToBiffEscapement(key.Escapement));
            stream.WriteByte(key.UnderlineStyle);
            stream.WriteByte(key.Family);
            stream.WriteByte(key.CharacterSet);
            stream.WriteByte(0);
            stream.WriteByte(checked((byte)key.Name.Length));
            stream.WriteByte(flags);
            stream.Write(nameBytes, 0, nameBytes.Length);
            return stream.ToArray();
        }

        private static bool TryGetFontByteProperty(
            OpenXmlElement properties,
            string elementName,
            string context,
            string subject,
            byte defaultValue,
            out byte value,
            out string? reason) {
            value = defaultValue;
            reason = null;
            OpenXmlElement? element = properties.ChildElements.FirstOrDefault(child => IsSpreadsheetElement(child, elementName));
            if (element == null) {
                return true;
            }

            string text = element.GetAttribute("val", string.Empty).Value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text)) {
                return true;
            }

            if (!uint.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint parsed) || parsed > byte.MaxValue) {
                reason = $"{context} {subject} outside BIFF8 limits";
                return false;
            }

            value = checked((byte)parsed);
            return true;
        }

        private static bool TryGetUnderlineStyle(
            Underline? underline,
            string context,
            out byte underlineStyle,
            out string? reason) {
            underlineStyle = 0;
            reason = null;
            if (underline == null) {
                return true;
            }

            UnderlineValues value = underline.Val?.Value ?? UnderlineValues.Single;
            if (value == UnderlineValues.None) {
                return true;
            }

            if (value == UnderlineValues.Single) {
                underlineStyle = 0x01;
                return true;
            }

            if (value == UnderlineValues.Double) {
                underlineStyle = 0x02;
                return true;
            }

            if (value == UnderlineValues.SingleAccounting) {
                underlineStyle = 0x21;
                return true;
            }

            if (value == UnderlineValues.DoubleAccounting) {
                underlineStyle = 0x22;
                return true;
            }

            reason = $"{context} underline styles outside BIFF8 limits";
            return false;
        }

        private static bool TryGetFontEscapement(
            VerticalTextAlignment? verticalTextAlignment,
            string context,
            out LegacyXlsFontEscapement escapement,
            out string? reason) {
            escapement = LegacyXlsFontEscapement.None;
            reason = null;
            VerticalAlignmentRunValues? value = verticalTextAlignment?.Val?.Value;
            if (!value.HasValue) {
                return true;
            }

            if (value == VerticalAlignmentRunValues.Superscript) {
                escapement = LegacyXlsFontEscapement.Superscript;
                return true;
            }

            if (value == VerticalAlignmentRunValues.Subscript) {
                escapement = LegacyXlsFontEscapement.Subscript;
                return true;
            }

            if (value == VerticalAlignmentRunValues.Baseline) {
                return true;
            }

            reason = $"{context} font vertical text alignment outside BIFF8 limits";
            return false;
        }

        private static bool IsSpreadsheetElement(OpenXmlElement element, string localName) {
            return string.Equals(element.LocalName, localName, StringComparison.Ordinal)
                && string.Equals(element.NamespaceUri, "http://schemas.openxmlformats.org/spreadsheetml/2006/main", StringComparison.Ordinal);
        }

        private static ushort BuildOptions(FontKey key) {
            ushort options = 0;
            if (key.Italic) {
                options |= 0x0002;
            }

            if (key.Strikeout) {
                options |= 0x0008;
            }

            if (key.Outline) {
                options |= 0x0010;
            }

            if (key.Shadow) {
                options |= 0x0020;
            }

            if (key.Condense) {
                options |= 0x0040;
            }

            if (key.Extend) {
                options |= 0x0080;
            }

            return options;
        }

        private static ushort ToBiffEscapement(LegacyXlsFontEscapement escapement) {
            return escapement == LegacyXlsFontEscapement.Superscript
                ? (ushort)1
                : escapement == LegacyXlsFontEscapement.Subscript
                    ? (ushort)2
                    : (ushort)0;
        }

        private static byte[] BuildPalettePayload(IReadOnlyList<string> paletteColors) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, checked((ushort)paletteColors.Count));
            foreach (string argb in paletteColors) {
                stream.WriteByte(byte.Parse(argb.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
                stream.WriteByte(byte.Parse(argb.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
                stream.WriteByte(byte.Parse(argb.Substring(6, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
                stream.WriteByte(0);
            }

            return stream.ToArray();
        }

        private static byte[] EncodeUnicodeString(string text, out byte flags) {
            if (CanUseCompressedString(text)) {
                flags = 0;
                return Encoding.ASCII.GetBytes(text);
            }

            flags = 1;
            return Encoding.Unicode.GetBytes(text);
        }

        private static bool CanUseCompressedString(string text) {
            for (int i = 0; i < text.Length; i++) {
                if (text[i] > 0x7f) {
                    return false;
                }
            }

            return true;
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
        }

        private readonly struct FontKey : IEquatable<FontKey> {
            internal FontKey(string name, int sizeTwips, ushort colorIndex, bool bold, bool italic, byte underlineStyle, bool strikeout, bool outline, bool shadow, bool condense, bool extend, LegacyXlsFontEscapement escapement, byte family, byte characterSet) {
                Name = name;
                SizeTwips = sizeTwips;
                ColorIndex = colorIndex;
                Bold = bold;
                Italic = italic;
                UnderlineStyle = underlineStyle;
                Strikeout = strikeout;
                Outline = outline;
                Shadow = shadow;
                Condense = condense;
                Extend = extend;
                Escapement = escapement;
                Family = family;
                CharacterSet = characterSet;
            }

            internal string Name { get; }
            internal int SizeTwips { get; }
            internal ushort ColorIndex { get; }
            internal bool Bold { get; }
            internal bool Italic { get; }
            internal byte UnderlineStyle { get; }
            internal bool Strikeout { get; }
            internal bool Outline { get; }
            internal bool Shadow { get; }
            internal bool Condense { get; }
            internal bool Extend { get; }
            internal LegacyXlsFontEscapement Escapement { get; }
            internal byte Family { get; }
            internal byte CharacterSet { get; }

            public bool Equals(FontKey other) {
                return string.Equals(Name, other.Name, StringComparison.OrdinalIgnoreCase)
                    && SizeTwips == other.SizeTwips
                    && ColorIndex == other.ColorIndex
                    && Bold == other.Bold
                    && Italic == other.Italic
                    && UnderlineStyle == other.UnderlineStyle
                    && Strikeout == other.Strikeout
                    && Outline == other.Outline
                    && Shadow == other.Shadow
                    && Condense == other.Condense
                    && Extend == other.Extend
                    && Escapement == other.Escapement
                    && Family == other.Family
                    && CharacterSet == other.CharacterSet;
            }

            public override bool Equals(object? obj) {
                return obj is FontKey other && Equals(other);
            }

            public override int GetHashCode() {
                unchecked {
                    int hash = StringComparer.OrdinalIgnoreCase.GetHashCode(Name);
                    hash = (hash * 397) ^ SizeTwips;
                    hash = (hash * 397) ^ ColorIndex.GetHashCode();
                    hash = (hash * 397) ^ Bold.GetHashCode();
                    hash = (hash * 397) ^ Italic.GetHashCode();
                    hash = (hash * 397) ^ UnderlineStyle.GetHashCode();
                    hash = (hash * 397) ^ Strikeout.GetHashCode();
                    hash = (hash * 397) ^ Outline.GetHashCode();
                    hash = (hash * 397) ^ Shadow.GetHashCode();
                    hash = (hash * 397) ^ Condense.GetHashCode();
                    hash = (hash * 397) ^ Extend.GetHashCode();
                    hash = (hash * 397) ^ Escapement.GetHashCode();
                    hash = (hash * 397) ^ Family.GetHashCode();
                    hash = (hash * 397) ^ CharacterSet.GetHashCode();
                    return hash;
                }
            }
        }
    }
}
