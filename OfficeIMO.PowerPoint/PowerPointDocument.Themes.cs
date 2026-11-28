using System.Xml.Linq;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private const int ThemeNamespace = OfficeXmlNamespaces.DrawingML;

        /// <summary>
        ///     Gets or sets the name of the presentation theme.
        /// </summary>
        public string ThemeName {
            get {
                ThrowIfDisposed();
                return RequireMasterPartRef(0).RelatedPart(PowerPointPartRef.ThemeRelationshipType).Root.Attribute("name") ?? string.Empty;
            }
            set {
                ThrowIfDisposed();
                PowerPointPartRef theme = EnsureThemePart(RequireMasterPartRef(0));
                theme.Root.SetAttribute("name", value);
                theme.MarkEdited();
            }
        }

        /// <summary>
        ///     Gets a theme color value in hex format (e.g. "FF0000").
        /// </summary>
        public string? GetThemeColor(PowerPointThemeColor color, int masterIndex = 0) {
            ThrowIfDisposed();
            OfficeXmlNode scheme = ThemeElements(RequireMasterPartRef(masterIndex)).Child(ThemeNamespace, "clrScheme");
            return scheme.IsNone ? null : GetThemeColorValue(scheme.Child(ThemeNamespace, ColorElementName(color)));
        }

        /// <summary>
        ///     Sets a theme color value in hex format (e.g. "FF0000").
        /// </summary>
        public void SetThemeColor(PowerPointThemeColor color, string hexValue, int masterIndex = 0) {
            ThrowIfDisposed();
            ArgumentException.ThrowIfNullOrWhiteSpace(hexValue);

            PowerPointPartRef theme = EnsureThemePart(RequireMasterPartRef(masterIndex));
            OfficeXmlNode scheme = EnsureScheme(theme, "clrScheme");
            string name = ColorElementName(color);
            OfficeXmlNode element = scheme.Child(ThemeNamespace, name);
            // A missing slot goes last, as the model added it.
            if (element.IsNone) element = scheme.Append(scheme.CreateElement(ThemeNamespace, name));
            element.RemoveChildren(ThemeNamespace, "srgbClr");
            element.RemoveChildren(ThemeNamespace, "sysClr");
            element.Append(element.CreateElement(ThemeNamespace, "srgbClr")).SetAttribute("val", hexValue);
            theme.MarkEdited();
            RefreshLayoutTheme(masterIndex);
        }

        /// <summary>
        ///     Sets multiple theme colors at once.
        /// </summary>
        public void SetThemeColors(IDictionary<PowerPointThemeColor, string> colors, int masterIndex = 0) {
            ThrowIfDisposed();
            ArgumentNullException.ThrowIfNull(colors);

            foreach (KeyValuePair<PowerPointThemeColor, string> entry in colors) {
                SetThemeColor(entry.Key, entry.Value, masterIndex);
            }
        }

        /// <summary>
        ///     Sets a theme color value across all masters.
        /// </summary>
        public void SetThemeColorForAllMasters(PowerPointThemeColor color, string hexValue) {
            ThrowIfDisposed();
            int masterCount = MasterCount;
            for (int i = 0; i < masterCount; i++) {
                SetThemeColor(color, hexValue, i);
            }
        }

        /// <summary>
        ///     Sets multiple theme colors across all masters.
        /// </summary>
        public void SetThemeColorsForAllMasters(IDictionary<PowerPointThemeColor, string> colors) {
            ThrowIfDisposed();
            ArgumentNullException.ThrowIfNull(colors);
            int masterCount = MasterCount;
            for (int i = 0; i < masterCount; i++) {
                SetThemeColors(colors, i);
            }
        }

        /// <summary>
        ///     Returns the theme colors that are defined on the master.
        /// </summary>
        public IReadOnlyDictionary<PowerPointThemeColor, string> GetThemeColors(int masterIndex = 0) {
            ThrowIfDisposed();
            OfficeXmlNode scheme = ThemeElements(RequireMasterPartRef(masterIndex)).Child(ThemeNamespace, "clrScheme");
            var colors = new Dictionary<PowerPointThemeColor, string>();
            if (scheme.IsNone) {
                return colors;
            }

            foreach (PowerPointThemeColor color in global::OfficeIMO.Internal.EnumCompat.GetValues<PowerPointThemeColor>()) {
                string? hexValue = GetThemeColorValue(scheme.Child(ThemeNamespace, ColorElementName(color)));
                if (!string.IsNullOrEmpty(hexValue)) {
                    colors[color] = hexValue!;
                }
            }

            return colors;
        }

        /// <summary>
        ///     Gets the major/minor Latin fonts for the theme.
        /// </summary>
        public PowerPointThemeFontInfo GetThemeLatinFonts(int masterIndex = 0) {
            ThrowIfDisposed();
            PowerPointThemeFontSet fonts = GetThemeFonts(masterIndex);
            return new PowerPointThemeFontInfo(fonts.MajorLatin, fonts.MinorLatin);
        }

        /// <summary>
        ///     Sets the major/minor Latin fonts for the theme.
        /// </summary>
        public void SetThemeLatinFonts(string majorLatin, string minorLatin, int masterIndex = 0) {
            ThrowIfDisposed();
            ArgumentException.ThrowIfNullOrWhiteSpace(majorLatin);
            ArgumentException.ThrowIfNullOrWhiteSpace(minorLatin);

            SetThemeFonts(new PowerPointThemeFontSet(majorLatin, minorLatin, null, null, null, null),
                masterIndex, keepExistingWhenNull: true);
        }

        /// <summary>
        ///     Sets the major/minor Latin fonts across all masters.
        /// </summary>
        public void SetThemeLatinFontsForAllMasters(string majorLatin, string minorLatin) {
            ThrowIfDisposed();
            int masterCount = MasterCount;
            for (int i = 0; i < masterCount; i++) {
                SetThemeLatinFonts(majorLatin, minorLatin, i);
            }
        }

        /// <summary>
        ///     Gets the major/minor fonts (Latin, East Asian, and complex script).
        /// </summary>
        public PowerPointThemeFontSet GetThemeFonts(int masterIndex = 0) {
            ThrowIfDisposed();
            OfficeXmlNode scheme = ThemeElements(RequireMasterPartRef(masterIndex)).Child(ThemeNamespace, "fontScheme");
            OfficeXmlNode major = scheme.Child(ThemeNamespace, "majorFont"), minor = scheme.Child(ThemeNamespace, "minorFont");

            return new PowerPointThemeFontSet(
                major.Child(ThemeNamespace, "latin").Attribute("typeface"),
                minor.Child(ThemeNamespace, "latin").Attribute("typeface"),
                major.Child(ThemeNamespace, "ea").Attribute("typeface"),
                minor.Child(ThemeNamespace, "ea").Attribute("typeface"),
                major.Child(ThemeNamespace, "cs").Attribute("typeface"),
                minor.Child(ThemeNamespace, "cs").Attribute("typeface"));
        }

        /// <summary>
        ///     Sets the major/minor fonts (Latin, East Asian, and complex script).
        /// </summary>
        public void SetThemeFonts(PowerPointThemeFontSet fonts, int masterIndex = 0, bool keepExistingWhenNull = true) {
            ThrowIfDisposed();
            PowerPointPartRef theme = EnsureThemePart(RequireMasterPartRef(masterIndex));
            OfficeXmlNode scheme = EnsureScheme(theme, "fontScheme");
            OfficeXmlNode major = scheme.GetOrAddChild(ThemeNamespace, "majorFont", AFontSchemeView.Order);
            OfficeXmlNode minor = scheme.GetOrAddChild(ThemeNamespace, "minorFont", AFontSchemeView.Order);

            SetThemeFont(major, fonts.MajorLatin, fonts.MajorEastAsian, fonts.MajorComplexScript, keepExistingWhenNull);
            SetThemeFont(minor, fonts.MinorLatin, fonts.MinorEastAsian, fonts.MinorComplexScript, keepExistingWhenNull);
            theme.MarkEdited();
            RefreshLayoutTheme(masterIndex);
        }

        /// <summary>
        ///     Sets the major/minor fonts (Latin, East Asian, and complex script) across all masters.
        /// </summary>
        public void SetThemeFontsForAllMasters(PowerPointThemeFontSet fonts, bool keepExistingWhenNull = true) {
            ThrowIfDisposed();
            int masterCount = MasterCount;
            for (int i = 0; i < masterCount; i++) {
                SetThemeFonts(fonts, i, keepExistingWhenNull);
            }
        }

        /// <summary>
        ///     Sets the theme name across all masters.
        /// </summary>
        public void SetThemeNameForAllMasters(string name) {
            ThrowIfDisposed();
            ArgumentException.ThrowIfNullOrWhiteSpace(name);
            int masterCount = MasterCount;
            for (int i = 0; i < masterCount; i++) {
                PowerPointPartRef theme = EnsureThemePart(RequireMasterPartRef(i));
                theme.Root.SetAttribute("name", name);
                theme.MarkEdited();
            }
        }

        /// <summary>Sets the name of a selected master's theme without restyling its layouts.</summary>
        /// <param name="name">Non-empty theme name.</param>
        /// <param name="masterIndex">Zero-based master position.</param>
        public void SetThemeName(string name, int masterIndex = 0) {
            ThrowIfDisposed();
            ArgumentException.ThrowIfNullOrWhiteSpace(name);
            PowerPointPartRef theme = EnsureThemePart(RequireMasterPartRef(masterIndex));
            OfficeXmlNode root = theme.Root;
            root.GetOrAddChild(ThemeNamespace, "themeElements", AOfficeStyleSheetView.Order);
            root.SetAttribute("name", name);
            theme.MarkEdited();
        }

        /// <summary>
        ///     Gets the list of table styles available in the presentation.
        /// </summary>
        public IReadOnlyList<PowerPointTableStyleInfo> TableStyles {
            get {
                ThrowIfDisposed();
                List<PowerPointTableStyleInfo> styles = new();
                HashSet<string> seenStyleIds = new(StringComparer.OrdinalIgnoreCase);
                foreach (OfficeXmlNode style in EnsureTableStylesPart().Root.Children(ThemeNamespace, "tblStyle")) {
                    string styleId = style.Attribute("styleId") ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(styleId) || !seenStyleIds.Add(styleId)) {
                        continue;
                    }

                    styles.Add(new PowerPointTableStyleInfo(styleId, style.Attribute("styleName") ?? string.Empty));
                }

                if (styles.Count == 0) {
                    using Stream? resource = typeof(PowerPointDocument).Assembly
                        .GetManifestResourceStream(TableStylesResourceName);
                    if (resource != null) {
                        AddTableStylesFromStream(styles, seenStyleIds, resource);
                    }
                }

                return styles.Count == 0 ? Array.Empty<PowerPointTableStyleInfo>() : styles;
            }
        }

        private static void AddTableStylesFromStream(List<PowerPointTableStyleInfo> styles, ISet<string> seenStyleIds, Stream stream) {
            XDocument document = PowerPointXmlReader.LoadPackagePartXml(stream);
            XElement? root = document.Root;
            if (root == null) {
                return;
            }

            XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
            foreach (XElement style in root.Elements(drawing + "tblStyle")) {
                string styleId = style.Attribute("styleId")?.Value ?? string.Empty;
                if (string.IsNullOrWhiteSpace(styleId) || !seenStyleIds.Add(styleId)) {
                    continue;
                }

                string name = style.Attribute("styleName")?.Value ?? string.Empty;
                styles.Add(new PowerPointTableStyleInfo(styleId, name));
            }
        }

        // The number of slide masters, in relationship order as the index-based APIs count them.
        private int MasterCount => PresentationPartRef.RelatedParts(PowerPointPartRef.SlideMasterRelationshipType).Count();

        // The theme elements of a master's theme, or none.
        private static OfficeXmlNode ThemeElements(PowerPointPartRef master) =>
            master.RelatedPart(PowerPointPartRef.ThemeRelationshipType).Root.Child(ThemeNamespace, "themeElements");

        // A master's theme part with its theme elements; a master without a theme gets one, as the model adds it.
        private PowerPointPartRef EnsureThemePart(PowerPointPartRef master) {
            PowerPointPartRef theme = master.RelatedPart(PowerPointPartRef.ThemeRelationshipType);
            if (theme.IsNone || theme.Root.IsNone) theme = PowerPointUtils.EnsureThemePart(this, master);
            theme.Root.GetOrAddChild(ThemeNamespace, "themeElements", AOfficeStyleSheetView.Order);
            return theme;
        }

        // The colour or font scheme of a theme, added with the name Office when missing.
        private static OfficeXmlNode EnsureScheme(PowerPointPartRef theme, string localName) {
            OfficeXmlNode elements = theme.Root.Child(ThemeNamespace, "themeElements");
            OfficeXmlNode scheme = elements.Child(ThemeNamespace, localName);
            if (!scheme.IsNone) return scheme;
            scheme = elements.AddOrdered(ThemeNamespace, localName, ABaseStylesView.Order);
            scheme.SetAttribute("name", "Office");
            return scheme;
        }

        private static string? GetThemeColorValue(OfficeXmlNode element) {
            string? rgbColor = element.Child(ThemeNamespace, "srgbClr").Attribute("val");
            if (!string.IsNullOrWhiteSpace(rgbColor)) {
                return rgbColor;
            }

            return element.Child(ThemeNamespace, "sysClr").Attribute("lastClr");
        }

        // Replaces the Latin, East Asian and complex script fonts, first in the collection in that order.
        private static void SetThemeFont(OfficeXmlNode parent, string? latin, string? eastAsian, string? complexScript, bool keepExistingWhenNull) {
            OfficeXmlNode resolvedLatin = ResolveThemeFont(parent, "latin", latin, keepExistingWhenNull);
            OfficeXmlNode resolvedEastAsian = ResolveThemeFont(parent, "ea", eastAsian, keepExistingWhenNull);
            OfficeXmlNode resolvedComplexScript = ResolveThemeFont(parent, "cs", complexScript, keepExistingWhenNull);

            parent.RemoveChildren(ThemeNamespace, "latin");
            parent.RemoveChildren(ThemeNamespace, "ea");
            parent.RemoveChildren(ThemeNamespace, "cs");

            OfficeXmlNode first = parent.FirstChild;
            parent.InsertBefore(resolvedLatin, first);
            parent.InsertBefore(resolvedEastAsian, first);
            parent.InsertBefore(resolvedComplexScript, first);
        }

        private static OfficeXmlNode ResolveThemeFont(OfficeXmlNode parent, string localName, string? typeface, bool keepExistingWhenNull) {
            OfficeXmlNode existing = parent.Child(ThemeNamespace, localName);
            existing = existing.IsNone ? default : existing.Clone();
            if (typeface == null) {
                // DrawingML requires all three font slots even when no typeface is specified.
                if (keepExistingWhenNull && !existing.IsNone) return existing;
                OfficeXmlNode empty = parent.CreateElement(ThemeNamespace, localName);
                empty.SetAttribute("typeface", string.Empty);
                return empty;
            }
            ArgumentException.ThrowIfNullOrWhiteSpace(typeface);

            OfficeXmlNode font = existing.IsNone ? parent.CreateElement(ThemeNamespace, localName) : existing;
            if (localName == "latin") {
                if (!string.Equals(font.Attribute("typeface"), typeface, StringComparison.Ordinal) || font.Attribute("panose") == null) {
                    PowerPointFontClassification.Apply(font, typeface);
                }
            } else {
                font.SetAttribute("typeface", typeface);
                if (!existing.IsNone && !string.Equals(parent.Child(ThemeNamespace, localName).Attribute("typeface"), typeface, StringComparison.Ordinal)) {
                    font.RemoveAttribute("panose");
                }
            }
            return font;
        }

        private static string ColorElementName(PowerPointThemeColor color) => color switch {
            PowerPointThemeColor.Dark1 => "dk1",
            PowerPointThemeColor.Light1 => "lt1",
            PowerPointThemeColor.Dark2 => "dk2",
            PowerPointThemeColor.Light2 => "lt2",
            PowerPointThemeColor.Accent1 => "accent1",
            PowerPointThemeColor.Accent2 => "accent2",
            PowerPointThemeColor.Accent3 => "accent3",
            PowerPointThemeColor.Accent4 => "accent4",
            PowerPointThemeColor.Accent5 => "accent5",
            PowerPointThemeColor.Accent6 => "accent6",
            PowerPointThemeColor.Hyperlink => "hlink",
            PowerPointThemeColor.FollowedHyperlink => "folHlink",
            _ => "dk1"
        };
    }
}
