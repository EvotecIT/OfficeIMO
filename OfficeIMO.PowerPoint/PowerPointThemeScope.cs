using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        /// <summary>
        /// The presentation part: the package's main part while the SDK is closed, otherwise the SDK's.
        /// </summary>
        internal PowerPointPartRef PresentationPartRef {
            get {
                if (!IsSdkLoaded && ReadablePackage is { } package) {
                    OfficeXmlTree relationships = package.GetRelationships(null);
                    for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                        node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                        string? type = relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Type");
                        if (type == null || OfficeXmlNamespaces.NormalizeRelationshipType(type) != OfficeOpcPackage.OfficeDocumentRelationshipType) continue;
                        string? main = OfficeOpcPartNames.ResolveRelationshipTarget("/_rels/.rels", relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
                        if (main != null && package.PartExists(main)) return PowerPointPartRef.Package(this, main);
                        break;
                    }
                }
                return SdkPresentationPartRef;
            }
        }

        /// <summary>The Latin heading and body typefaces the first master's theme part names now.</summary>
        internal (string? Heading, string? Body) ThemeFontFamilies() {
            OfficeXmlNode fonts = MasterPartRef(0).RelatedPart(PowerPointPartRef.ThemeRelationshipType).Root
                .Child(OfficeXmlNamespaces.DrawingML, "themeElements").Child(OfficeXmlNamespaces.DrawingML, "fontScheme");
            string? heading = LatinTypeface(fonts, "majorFont");
            string? body = LatinTypeface(fonts, "minorFont");
            return (string.IsNullOrWhiteSpace(heading) ? null : heading, string.IsNullOrWhiteSpace(body) ? null : body);
        }

        /// <summary>Reads a master's current colours and fonts without restyling its native layouts.</summary>
        internal PowerPointLayoutTheme ReadLayoutTheme(int masterIndex = 0) => ReadLayoutTheme(MasterPartRef(masterIndex), default, default);

        /// <summary>
        /// Reads a master's current colours and fonts, with the colour and font schemes a theme override puts in their
        /// place when given.
        /// </summary>
        internal PowerPointLayoutTheme ReadLayoutTheme(PowerPointPartRef master, OfficeXmlNode colorsOverride, OfficeXmlNode fontsOverride) {
            OfficeTheme theme = _layoutTheme?.Theme ?? _baseTheme ?? OfficeTheme.Default;
            OfficeXmlNode nativeTheme = master.IsNone ? default : master.RelatedPart(PowerPointPartRef.ThemeRelationshipType).Root;
            OfficeXmlNode elements = nativeTheme.Child(OfficeXmlNamespaces.DrawingML, "themeElements");
            if (elements.IsNone && colorsOverride.IsNone && fontsOverride.IsNone) return PowerPointLayoutTheme.From(theme);

            OfficeXmlNode colors = colorsOverride.IsNone ? elements.Child(OfficeXmlNamespaces.DrawingML, "clrScheme") : colorsOverride;
            OfficeColor? Read(string slot) {
                OfficeXmlNode container = colors.Child(OfficeXmlNamespaces.DrawingML, slot);
                string? hex = container.Child(OfficeXmlNamespaces.DrawingML, "srgbClr").Attribute("val")
                              ?? container.Child(OfficeXmlNamespaces.DrawingML, "sysClr").Attribute("lastClr");
                if (string.IsNullOrWhiteSpace(hex) || hex!.Length != 6) return null;
                try {
                    return OfficeColor.ParseHex(hex);
                } catch (FormatException) {
                    return null;
                }
            }

            OfficeThemePalette palette = theme.Palette.With(
                background: Read("lt1"),
                text: Read("dk1"),
                textMuted: Read("dk2"),
                accent1: Read("accent1"),
                accent2: Read("accent2"),
                accent3: Read("accent3"),
                accent4: Read("accent4"),
                accent5: Read("accent5"),
                accent6: Read("accent6"));
            OfficeXmlNode fontScheme = fontsOverride.IsNone ? elements.Child(OfficeXmlNamespaces.DrawingML, "fontScheme") : fontsOverride;
            string? heading = LatinTypeface(fontScheme, "majorFont");
            string? body = LatinTypeface(fontScheme, "minorFont");
            OfficeThemeFonts fonts = new(
                string.IsNullOrWhiteSpace(heading) ? theme.Fonts.Heading : theme.Fonts.Heading.With(family: heading),
                string.IsNullOrWhiteSpace(body) ? theme.Fonts.Body : theme.Fonts.Body.With(family: body));
            string? name = nativeTheme.Attribute("name");
            return PowerPointLayoutTheme.From(theme.WithName(string.IsNullOrWhiteSpace(name) ? theme.Name : name!)
                .WithPalette(palette).WithFonts(fonts)
                .WithSurfaces(theme.Surfaces.With(surface: Read("lt2"))));
        }

        private static string? LatinTypeface(OfficeXmlNode fontScheme, string font) =>
            fontScheme.Child(OfficeXmlNamespaces.DrawingML, font).Child(OfficeXmlNamespaces.DrawingML, "latin").Attribute("typeface");

        /// <summary>The slide master at <paramref name="masterIndex"/> in relationship order, or none.</summary>
        internal PowerPointPartRef MasterPartRef(int masterIndex) =>
            masterIndex < 0 ? default : PresentationPartRef.RelatedParts(PowerPointPartRef.SlideMasterRelationshipType).Skip(masterIndex).FirstOrDefault();
    }

    /// <summary>
    /// The colour, font and format schemes in effect for a slide, layout or master part: a slide's theme override, then
    /// its layout's, then the theme of the master. Read through <see cref="PowerPointPartRef"/>, so it works on the
    /// package and on the SDK alike.
    /// </summary>
    internal readonly struct PowerPointThemeScope {
        private PowerPointThemeScope(PowerPointPartRef master, OfficeXmlNode theme, OfficeXmlNode colors, OfficeXmlNode fonts, OfficeXmlNode format, OfficeXmlNode colorMap) {
            Master = master;
            ColorMap = colorMap;
            Theme = theme;
            ColorScheme = colors;
            FontScheme = fonts;
            FormatScheme = format;
        }

        /// <summary>The master part, or none when the part has no master.</summary>
        internal PowerPointPartRef Master { get; }

        /// <summary>The master theme's root element.</summary>
        internal OfficeXmlNode Theme { get; }

        internal OfficeXmlNode ColorScheme { get; }

        internal OfficeXmlNode FontScheme { get; }

        internal OfficeXmlNode FormatScheme { get; }

        /// <summary>
        /// The color map scheme colors go through: the slide's override, else its layout's, else the master's.
        /// </summary>
        internal OfficeXmlNode ColorMap { get; }

        internal static PowerPointThemeScope For(PowerPointPartRef part) {
            PowerPointPartRef slide = default, layout = default, master = default;
            switch (part.Kind) {
                case PowerPointPartKind.Slide:
                    slide = part;
                    layout = part.RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType);
                    break;
                case PowerPointPartKind.SlideLayout:
                    layout = part;
                    break;
                case PowerPointPartKind.SlideMaster:
                    master = part;
                    break;
                default:
                    return default;
            }
            if (!layout.IsNone) master = layout.RelatedPart(PowerPointPartRef.SlideMasterRelationshipType);
            OfficeXmlNode theme = master.IsNone ? default : master.RelatedPart(PowerPointPartRef.ThemeRelationshipType).Root;
            OfficeXmlNode elements = theme.Child(OfficeXmlNamespaces.DrawingML, "themeElements");
            OfficeXmlNode slideOverride = Override(slide), layoutOverride = Override(layout);
            return new PowerPointThemeScope(master, theme,
                First(slideOverride, layoutOverride, elements, "clrScheme"),
                First(slideOverride, layoutOverride, elements, "fontScheme"),
                First(slideOverride, layoutOverride, elements, "fmtScheme"),
                ColorMapFor(slide, layout, master));
        }

        private static OfficeXmlNode ColorMapFor(PowerPointPartRef slide, PowerPointPartRef layout, PowerPointPartRef master) {
            OfficeXmlNode masterMap = master.IsNone ? default : master.Root.Child(OfficeXmlNamespaces.PresentationML, "clrMap");
            OfficeXmlNode slideOverride = slide.IsNone ? default : slide.Root.Child(OfficeXmlNamespaces.PresentationML, "clrMapOvr");
            if (!slideOverride.Child(OfficeXmlNamespaces.DrawingML, "masterClrMapping").IsNone) return masterMap;
            OfficeXmlNode slideMap = slideOverride.Child(OfficeXmlNamespaces.DrawingML, "overrideClrMapping");
            if (!slideMap.IsNone) return slideMap;
            OfficeXmlNode layoutMap = layout.IsNone ? default
                : layout.Root.Child(OfficeXmlNamespaces.PresentationML, "clrMapOvr").Child(OfficeXmlNamespaces.DrawingML, "overrideClrMapping");
            return layoutMap.IsNone ? masterMap : layoutMap;
        }

        private static OfficeXmlNode Override(PowerPointPartRef part) =>
            part.IsNone ? default : part.RelatedPart(PowerPointPartRef.ThemeOverrideRelationshipType).Root;

        private static OfficeXmlNode First(OfficeXmlNode slideOverride, OfficeXmlNode layoutOverride, OfficeXmlNode elements, string name) {
            OfficeXmlNode scheme = slideOverride.Child(OfficeXmlNamespaces.DrawingML, name);
            if (scheme.IsNone) scheme = layoutOverride.Child(OfficeXmlNamespaces.DrawingML, name);
            return scheme.IsNone ? elements.Child(OfficeXmlNamespaces.DrawingML, name) : scheme;
        }
    }
}
