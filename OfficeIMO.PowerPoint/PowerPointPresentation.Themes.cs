using System;
using System.IO;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        /// <summary>
        ///     Gets or sets the name of the presentation theme.
        /// </summary>
        public string ThemeName {
            get {
                ThrowIfDisposed();
                SlideMasterPart master = _presentationPart.SlideMasterParts.First();
                return master.ThemePart?.Theme?.Name?.Value ?? string.Empty;
            }
            set {
                ThrowIfDisposed();
                SlideMasterPart master = _presentationPart.SlideMasterParts.First();
                ThemePart themePart = master.ThemePart ?? master.AddNewPart<ThemePart>();
                if (themePart.Theme == null) {
                    themePart.Theme = new A.Theme { ThemeElements = new A.ThemeElements() };
                }

                themePart.Theme.Name = value;
            }
        }

        /// <summary>
        ///     Gets a theme color value in hex format (e.g. "FF0000").
        /// </summary>
        public string? GetThemeColor(PowerPointThemeColor color, int masterIndex = 0) {
            ThrowIfDisposed();
            SlideMasterPart masterPart = GetSlideMasterPart(masterIndex);
            A.ColorScheme? scheme = masterPart.ThemePart?.Theme?.ThemeElements?.ColorScheme;
            if (scheme == null) {
                return null;
            }

            OpenXmlCompositeElement? element = GetColorElement(scheme, color);
            return GetThemeColorValue(element);
        }

        /// <summary>
        ///     Sets a theme color value in hex format (e.g. "FF0000").
        /// </summary>
        public void SetThemeColor(PowerPointThemeColor color, string hexValue, int masterIndex = 0) {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(hexValue)) {
                throw new ArgumentException("Theme color value cannot be null or empty.", nameof(hexValue));
            }

            SlideMasterPart masterPart = GetSlideMasterPart(masterIndex);
            A.ColorScheme scheme = EnsureColorScheme(masterPart);
            OpenXmlCompositeElement element = GetOrCreateColorElement(scheme, color);
            element.RemoveAllChildren<A.RgbColorModelHex>();
            element.RemoveAllChildren<A.SystemColor>();
            element.Append(new A.RgbColorModelHex { Val = hexValue });
        }

        /// <summary>
        ///     Sets multiple theme colors at once.
        /// </summary>
        public void SetThemeColors(IDictionary<PowerPointThemeColor, string> colors, int masterIndex = 0) {
            ThrowIfDisposed();
            if (colors == null) {
                throw new ArgumentNullException(nameof(colors));
            }

            foreach (KeyValuePair<PowerPointThemeColor, string> entry in colors) {
                SetThemeColor(entry.Key, entry.Value, masterIndex);
            }
        }

        /// <summary>
        ///     Sets a theme color value across all masters.
        /// </summary>
        public void SetThemeColorForAllMasters(PowerPointThemeColor color, string hexValue) {
            ThrowIfDisposed();
            int masterCount = _presentationPart.SlideMasterParts.Count();
            for (int i = 0; i < masterCount; i++) {
                SetThemeColor(color, hexValue, i);
            }
        }

        /// <summary>
        ///     Sets multiple theme colors across all masters.
        /// </summary>
        public void SetThemeColorsForAllMasters(IDictionary<PowerPointThemeColor, string> colors) {
            ThrowIfDisposed();
            if (colors == null) {
                throw new ArgumentNullException(nameof(colors));
            }
            int masterCount = _presentationPart.SlideMasterParts.Count();
            for (int i = 0; i < masterCount; i++) {
                SetThemeColors(colors, i);
            }
        }

        /// <summary>
        ///     Returns the theme colors that are defined on the master.        
        /// </summary>
        public IReadOnlyDictionary<PowerPointThemeColor, string> GetThemeColors(int masterIndex = 0) {
            ThrowIfDisposed();
            SlideMasterPart masterPart = GetSlideMasterPart(masterIndex);
            A.ColorScheme? scheme = masterPart.ThemePart?.Theme?.ThemeElements?.ColorScheme;
            if (scheme == null) {
                return new Dictionary<PowerPointThemeColor, string>();
            }

            var colors = new Dictionary<PowerPointThemeColor, string>();
            foreach (PowerPointThemeColor color in global::OfficeIMO.Internal.EnumCompat.GetValues<PowerPointThemeColor>()) {
                OpenXmlCompositeElement? element = GetColorElement(scheme, color);
                string? hexValue = GetThemeColorValue(element);
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
            if (string.IsNullOrWhiteSpace(majorLatin)) {
                throw new ArgumentException("Major font cannot be null or empty.", nameof(majorLatin));
            }
            if (string.IsNullOrWhiteSpace(minorLatin)) {
                throw new ArgumentException("Minor font cannot be null or empty.", nameof(minorLatin));
            }

            SetThemeFonts(new PowerPointThemeFontSet(majorLatin, minorLatin, null, null, null, null),
                masterIndex, keepExistingWhenNull: true);
        }

        /// <summary>
        ///     Sets the major/minor Latin fonts across all masters.
        /// </summary>
        public void SetThemeLatinFontsForAllMasters(string majorLatin, string minorLatin) {
            ThrowIfDisposed();
            int masterCount = _presentationPart.SlideMasterParts.Count();
            for (int i = 0; i < masterCount; i++) {
                SetThemeLatinFonts(majorLatin, minorLatin, i);
            }
        }

        /// <summary>
        ///     Gets the major/minor fonts (Latin, East Asian, and complex script).
        /// </summary>
        public PowerPointThemeFontSet GetThemeFonts(int masterIndex = 0) {
            ThrowIfDisposed();
            SlideMasterPart masterPart = GetSlideMasterPart(masterIndex);
            A.FontScheme? scheme = masterPart.ThemePart?.Theme?.ThemeElements?.FontScheme;

            return new PowerPointThemeFontSet(
                scheme?.MajorFont?.LatinFont?.Typeface,
                scheme?.MinorFont?.LatinFont?.Typeface,
                scheme?.MajorFont?.EastAsianFont?.Typeface,
                scheme?.MinorFont?.EastAsianFont?.Typeface,
                scheme?.MajorFont?.ComplexScriptFont?.Typeface,
                scheme?.MinorFont?.ComplexScriptFont?.Typeface);
        }

        /// <summary>
        ///     Sets the major/minor fonts (Latin, East Asian, and complex script).
        /// </summary>
        public void SetThemeFonts(PowerPointThemeFontSet fonts, int masterIndex = 0, bool keepExistingWhenNull = true) {
            ThrowIfDisposed();
            SlideMasterPart masterPart = GetSlideMasterPart(masterIndex);
            A.FontScheme scheme = EnsureFontScheme(masterPart);
            scheme.MajorFont ??= new A.MajorFont();
            scheme.MinorFont ??= new A.MinorFont();

            SetThemeFont(scheme.MajorFont, fonts.MajorLatin, fonts.MajorEastAsian, fonts.MajorComplexScript,
                keepExistingWhenNull);
            SetThemeFont(scheme.MinorFont, fonts.MinorLatin, fonts.MinorEastAsian, fonts.MinorComplexScript,
                keepExistingWhenNull);
        }

        /// <summary>
        ///     Sets the major/minor fonts (Latin, East Asian, and complex script) across all masters.
        /// </summary>
        public void SetThemeFontsForAllMasters(PowerPointThemeFontSet fonts, bool keepExistingWhenNull = true) {
            ThrowIfDisposed();
            int masterCount = _presentationPart.SlideMasterParts.Count();
            for (int i = 0; i < masterCount; i++) {
                SetThemeFonts(fonts, i, keepExistingWhenNull);
            }
        }

        /// <summary>
        ///     Sets the theme name across all masters.
        /// </summary>
        public void SetThemeNameForAllMasters(string name) {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Theme name cannot be null or empty.", nameof(name));
            }
            foreach (SlideMasterPart masterPart in _presentationPart.SlideMasterParts) {
                ThemePart themePart = masterPart.ThemePart ?? masterPart.AddNewPart<ThemePart>();
                themePart.Theme ??= new A.Theme { ThemeElements = new A.ThemeElements() };
                themePart.Theme.Name = name;
            }
        }

        /// <summary>
        ///     Gets the list of table styles available in the presentation.
        /// </summary>
        public IReadOnlyList<PowerPointTableStyleInfo> TableStyles {
            get {
                ThrowIfDisposed();
                TableStylesPart? stylesPart = _presentationPart.TableStylesPart;
                if (stylesPart?.TableStyleList == null) {
                    PowerPointUtils.CreateTableStylesPart(_presentationPart);
                    stylesPart = _presentationPart.TableStylesPart;
                }

                List<PowerPointTableStyleInfo> styles = new();
                HashSet<string> seenStyleIds = new(StringComparer.OrdinalIgnoreCase);
                A.TableStyleList? styleList = stylesPart?.TableStyleList;
                if (styleList != null) {
                    foreach (A.TableStyle style in styleList.Elements<A.TableStyle>()) {
                        string styleId = style.StyleId?.Value ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(styleId) || !seenStyleIds.Add(styleId)) {
                            continue;
                        }

                        string name = style.StyleName?.Value ?? string.Empty;
                        styles.Add(new PowerPointTableStyleInfo(styleId, name));
                    }
                }

                if (stylesPart != null) {
                    using Stream stream = stylesPart.GetStream(FileMode.Open, FileAccess.Read);
                    if (stream.Length > 0) {
                        AddTableStylesFromStream(styles, seenStyleIds, stream);
                    }
                }

                if (styles.Count == 0) {
                    using Stream? resource = typeof(PowerPointPresentation).Assembly
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

        private static string? GetThemeColorValue(OpenXmlCompositeElement? element) {
            string? rgbColor = element?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(rgbColor)) {
                return rgbColor;
            }

            return element?.GetFirstChild<A.SystemColor>()?.LastColor?.Value;
        }

        private static A.Theme EnsureTheme(SlideMasterPart masterPart) {
            ThemePart themePart = masterPart.ThemePart ?? masterPart.AddNewPart<ThemePart>();
            themePart.Theme ??= new A.Theme { ThemeElements = new A.ThemeElements() };
            themePart.Theme.ThemeElements ??= new A.ThemeElements();
            return themePart.Theme;
        }

        private static A.ColorScheme EnsureColorScheme(SlideMasterPart masterPart) {
            A.Theme theme = EnsureTheme(masterPart);
            theme.ThemeElements ??= new A.ThemeElements();
            A.ColorScheme scheme = theme.ThemeElements.ColorScheme ??= new A.ColorScheme { Name = "Office" };
            return scheme;
        }

        private static A.FontScheme EnsureFontScheme(SlideMasterPart masterPart) {
            A.Theme theme = EnsureTheme(masterPart);
            theme.ThemeElements ??= new A.ThemeElements();
            A.FontScheme scheme = theme.ThemeElements.FontScheme ??= new A.FontScheme { Name = "Office" };
            return scheme;
        }

        private static void SetThemeFont(OpenXmlCompositeElement parent, string? latin, string? eastAsian,
            string? complexScript, bool keepExistingWhenNull) {
            A.LatinFont? resolvedLatin = ResolveThemeFont<A.LatinFont>(parent, latin, keepExistingWhenNull);
            A.EastAsianFont? resolvedEastAsian =
                ResolveThemeFont<A.EastAsianFont>(parent, eastAsian, keepExistingWhenNull);
            A.ComplexScriptFont? resolvedComplexScript =
                ResolveThemeFont<A.ComplexScriptFont>(parent, complexScript, keepExistingWhenNull);

            parent.RemoveAllChildren<A.LatinFont>();
            parent.RemoveAllChildren<A.EastAsianFont>();
            parent.RemoveAllChildren<A.ComplexScriptFont>();

            int insertIndex = 0;
            if (resolvedLatin != null) {
                parent.InsertAt(resolvedLatin, insertIndex++);
            }
            if (resolvedEastAsian != null) {
                parent.InsertAt(resolvedEastAsian, insertIndex++);
            }
            if (resolvedComplexScript != null) {
                parent.InsertAt(resolvedComplexScript, insertIndex);
            }
        }

        private static TFont? ResolveThemeFont<TFont>(OpenXmlCompositeElement parent, string? typeface,
            bool keepExistingWhenNull) where TFont : A.TextFontType, new() {
            TFont? existing = parent.GetFirstChild<TFont>()?.CloneNode(true) as TFont;
            if (typeface == null) {
                return keepExistingWhenNull ? existing : null;
            }
            if (string.IsNullOrWhiteSpace(typeface)) {
                throw new ArgumentException("Font name cannot be null or empty.", nameof(typeface));
            }

            TFont font = existing ?? new TFont();
            font.Typeface = typeface;
            return font;
        }

        private static OpenXmlCompositeElement? GetColorElement(A.ColorScheme scheme, PowerPointThemeColor color) {
            return color switch {
                PowerPointThemeColor.Dark1 => scheme.GetFirstChild<A.Dark1Color>(),
                PowerPointThemeColor.Light1 => scheme.GetFirstChild<A.Light1Color>(),
                PowerPointThemeColor.Dark2 => scheme.GetFirstChild<A.Dark2Color>(),
                PowerPointThemeColor.Light2 => scheme.GetFirstChild<A.Light2Color>(),
                PowerPointThemeColor.Accent1 => scheme.GetFirstChild<A.Accent1Color>(),
                PowerPointThemeColor.Accent2 => scheme.GetFirstChild<A.Accent2Color>(),
                PowerPointThemeColor.Accent3 => scheme.GetFirstChild<A.Accent3Color>(),
                PowerPointThemeColor.Accent4 => scheme.GetFirstChild<A.Accent4Color>(),
                PowerPointThemeColor.Accent5 => scheme.GetFirstChild<A.Accent5Color>(),
                PowerPointThemeColor.Accent6 => scheme.GetFirstChild<A.Accent6Color>(),
                PowerPointThemeColor.Hyperlink => scheme.GetFirstChild<A.Hyperlink>(),
                PowerPointThemeColor.FollowedHyperlink => scheme.GetFirstChild<A.FollowedHyperlinkColor>(),
                _ => null
            };
        }

        private static OpenXmlCompositeElement GetOrCreateColorElement(A.ColorScheme scheme, PowerPointThemeColor color) {
            OpenXmlCompositeElement? element = GetColorElement(scheme, color);
            if (element != null) {
                return element;
            }

            element = color switch {
                PowerPointThemeColor.Dark1 => new A.Dark1Color(),
                PowerPointThemeColor.Light1 => new A.Light1Color(),
                PowerPointThemeColor.Dark2 => new A.Dark2Color(),
                PowerPointThemeColor.Light2 => new A.Light2Color(),
                PowerPointThemeColor.Accent1 => new A.Accent1Color(),
                PowerPointThemeColor.Accent2 => new A.Accent2Color(),
                PowerPointThemeColor.Accent3 => new A.Accent3Color(),
                PowerPointThemeColor.Accent4 => new A.Accent4Color(),
                PowerPointThemeColor.Accent5 => new A.Accent5Color(),
                PowerPointThemeColor.Accent6 => new A.Accent6Color(),
                PowerPointThemeColor.Hyperlink => new A.Hyperlink(),
                PowerPointThemeColor.FollowedHyperlink => new A.FollowedHyperlinkColor(),
                _ => new A.Dark1Color()
            };

            scheme.Append(element);
            return element;
        }

    }
}
