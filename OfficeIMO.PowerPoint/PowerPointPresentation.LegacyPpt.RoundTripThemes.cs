using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ApplyLegacyRoundTripTheme(
            SlideMasterPart target, LegacyPptRoundTripTheme? source) {
            if (source == null) return;
            if (!source.IsOverride && !string.IsNullOrWhiteSpace(
                    source.ThemeXml)) {
                ThemePart themePart = target.ThemePart
                    ?? target.AddNewPart<ThemePart>();
                themePart.Theme = new A.Theme(
                    NormalizeLegacyThemeRootXml(source.ThemeXml!));
                themePart.Theme.Save();
            }
            ApplyLegacyRoundTripColorMap(target.SlideMaster,
                source.ColorMappingXml);
        }

        private static void ApplyLegacyRoundTripTheme(
            NotesMasterPart target, LegacyPptRoundTripTheme? source) {
            if (source == null) return;
            if (!source.IsOverride && !string.IsNullOrWhiteSpace(
                    source.ThemeXml)) {
                ThemePart themePart = target.ThemePart
                    ?? target.AddNewPart<ThemePart>();
                themePart.Theme = new A.Theme(
                    NormalizeLegacyThemeRootXml(source.ThemeXml!));
                themePart.Theme.Save();
            }
            ApplyLegacyRoundTripColorMap(target.NotesMaster,
                source.ColorMappingXml);
        }

        private static void ApplyLegacyRoundTripTheme(
            HandoutMasterPart target, LegacyPptRoundTripTheme? source) {
            if (source == null) return;
            if (!source.IsOverride && !string.IsNullOrWhiteSpace(
                    source.ThemeXml)) {
                ThemePart themePart = target.ThemePart
                    ?? target.AddNewPart<ThemePart>();
                themePart.Theme = new A.Theme(
                    NormalizeLegacyThemeRootXml(source.ThemeXml!));
                themePart.Theme.Save();
            }
            ApplyLegacyRoundTripColorMap(target.HandoutMaster,
                source.ColorMappingXml);
        }

        private static void ApplyLegacyRoundTripTheme(
            OpenXmlPartContainer target, LegacyPptRoundTripTheme? source) {
            if (source == null) return;
            if (!string.IsNullOrWhiteSpace(source.ThemeXml)) {
                ThemeOverridePart themePart = target switch {
                    SlidePart slidePart => slidePart.ThemeOverridePart
                        ?? slidePart.AddNewPart<ThemeOverridePart>(),
                    SlideLayoutPart layoutPart => layoutPart.ThemeOverridePart
                        ?? layoutPart.AddNewPart<ThemeOverridePart>(),
                    NotesSlidePart notesPart => notesPart.ThemeOverridePart
                        ?? notesPart.AddNewPart<ThemeOverridePart>(),
                    _ => throw new ArgumentException(
                        "The target part cannot own a theme override.",
                        nameof(target))
                };
                themePart.ThemeOverride = CreateLegacyThemeOverride(source);
                themePart.ThemeOverride.Save();
            }

            P.ColorMapOverride? colorMap = CreateLegacyColorMapOverride(
                source.ColorMappingXml);
            if (colorMap == null) return;
            switch (target) {
                case SlidePart slidePart when slidePart.Slide != null:
                    slidePart.Slide.ColorMapOverride = colorMap;
                    break;
                case SlideLayoutPart layoutPart
                    when layoutPart.SlideLayout != null:
                    layoutPart.SlideLayout.ColorMapOverride = colorMap;
                    break;
                case NotesSlidePart notesPart when notesPart.NotesSlide != null:
                    notesPart.NotesSlide.ColorMapOverride = colorMap;
                    break;
            }
        }

        private static A.ThemeOverride CreateLegacyThemeOverride(
            LegacyPptRoundTripTheme source) {
            string xml = NormalizeLegacyThemeRootXml(source.ThemeXml!);
            if (source.IsOverride) return new A.ThemeOverride(xml);
            var theme = new A.Theme(xml);
            OpenXmlElement?[] candidates = {
                theme.ThemeElements?.ColorScheme?.CloneNode(true),
                theme.ThemeElements?.FontScheme?.CloneNode(true),
                theme.ThemeElements?.FormatScheme?.CloneNode(true)
            };
            OpenXmlElement[] elements = candidates.Where(element =>
                    element != null).Cast<OpenXmlElement>()
                .ToArray();
            return new A.ThemeOverride(elements);
        }

        private static void ApplyLegacyRoundTripColorMap(
            P.SlideMaster? target, string? xml) {
            if (target == null || !TryReadLegacyColorMap(xml,
                    out IReadOnlyDictionary<string, A.ColorSchemeIndexValues>
                        values)) return;
            P.ColorMap map = target.ColorMap ??= new P.ColorMap();
            ApplyLegacyColorMapValues(map, values);
        }

        private static void ApplyLegacyRoundTripColorMap(
            P.NotesMaster? target, string? xml) {
            if (target == null || !TryReadLegacyColorMap(xml,
                    out IReadOnlyDictionary<string, A.ColorSchemeIndexValues>
                        values)) return;
            P.ColorMap map = target.ColorMap ??= new P.ColorMap();
            ApplyLegacyColorMapValues(map, values);
        }

        private static void ApplyLegacyRoundTripColorMap(
            P.HandoutMaster? target, string? xml) {
            if (target == null || !TryReadLegacyColorMap(xml,
                    out IReadOnlyDictionary<string, A.ColorSchemeIndexValues>
                        values)) return;
            P.ColorMap map = target.ColorMap ??= new P.ColorMap();
            ApplyLegacyColorMapValues(map, values);
        }

        private static bool TryReadLegacyColorMap(string? xml,
            out IReadOnlyDictionary<string, A.ColorSchemeIndexValues> values) {
            var result = new Dictionary<string, A.ColorSchemeIndexValues>(
                StringComparer.Ordinal);
            values = result;
            if (string.IsNullOrWhiteSpace(xml)) return false;
            XElement root = XElement.Parse(xml, LoadOptions.None);
            if (root.Name.LocalName != "clrMap") return false;
            foreach (XAttribute attribute in root.Attributes().Where(
                         attribute => !attribute.IsNamespaceDeclaration)) {
                if (TryMapLegacyColorSchemeIndex(attribute.Value,
                        out A.ColorSchemeIndexValues value)) {
                    result[attribute.Name.LocalName] = value;
                }
            }
            return result.Count > 0;
        }

        private static string NormalizeLegacyThemeRootXml(string xml) =>
            XElement.Parse(xml, LoadOptions.PreserveWhitespace)
                .ToString(SaveOptions.DisableFormatting);

        private static P.ColorMapOverride? CreateLegacyColorMapOverride(
            string? xml) {
            if (string.IsNullOrWhiteSpace(xml)) return null;
            XElement root = XElement.Parse(xml, LoadOptions.None);
            XElement? mapping = root.Name.LocalName == "clrMap"
                ? ConvertToOverrideColorMapping(root)
                : root.Elements().FirstOrDefault(element =>
                    element.Name.LocalName is "masterClrMapping"
                        or "overrideClrMapping");
            if (mapping == null) return null;
            if (mapping.Name.LocalName == "masterClrMapping") {
                return new P.ColorMapOverride(new A.MasterColorMapping());
            }
            XElement normalized = ConvertToOverrideColorMapping(mapping);
            return new P.ColorMapOverride(new A.OverrideColorMapping(
                normalized.ToString(SaveOptions.DisableFormatting)));
        }

        private static XElement ConvertToOverrideColorMapping(
            XElement source) => new(
            XName.Get("overrideClrMapping",
                "http://schemas.openxmlformats.org/drawingml/2006/main"),
            new XAttribute(XNamespace.Xmlns + "a",
                "http://schemas.openxmlformats.org/drawingml/2006/main"),
            source.Attributes().Where(attribute =>
                    !attribute.IsNamespaceDeclaration)
                .Select(attribute => new XAttribute(
                    attribute.Name.LocalName, attribute.Value)));

        private static bool TryMapLegacyColorSchemeIndex(string value,
            out A.ColorSchemeIndexValues result) {
            switch (value) {
                case "dk1": result = A.ColorSchemeIndexValues.Dark1; return true;
                case "lt1": result = A.ColorSchemeIndexValues.Light1; return true;
                case "dk2": result = A.ColorSchemeIndexValues.Dark2; return true;
                case "lt2": result = A.ColorSchemeIndexValues.Light2; return true;
                case "accent1": result = A.ColorSchemeIndexValues.Accent1; return true;
                case "accent2": result = A.ColorSchemeIndexValues.Accent2; return true;
                case "accent3": result = A.ColorSchemeIndexValues.Accent3; return true;
                case "accent4": result = A.ColorSchemeIndexValues.Accent4; return true;
                case "accent5": result = A.ColorSchemeIndexValues.Accent5; return true;
                case "accent6": result = A.ColorSchemeIndexValues.Accent6; return true;
                case "hlink": result = A.ColorSchemeIndexValues.Hyperlink; return true;
                case "folHlink": result = A.ColorSchemeIndexValues.FollowedHyperlink; return true;
                default: result = default; return false;
            }
        }

        private static void ApplyLegacyColorMapValues(P.ColorMap target,
            IReadOnlyDictionary<string, A.ColorSchemeIndexValues> values) {
            if (values.TryGetValue("bg1", out A.ColorSchemeIndexValues bg1))
                target.Background1 = bg1;
            if (values.TryGetValue("tx1", out A.ColorSchemeIndexValues tx1))
                target.Text1 = tx1;
            if (values.TryGetValue("bg2", out A.ColorSchemeIndexValues bg2))
                target.Background2 = bg2;
            if (values.TryGetValue("tx2", out A.ColorSchemeIndexValues tx2))
                target.Text2 = tx2;
            if (values.TryGetValue("accent1", out A.ColorSchemeIndexValues accent1))
                target.Accent1 = accent1;
            if (values.TryGetValue("accent2", out A.ColorSchemeIndexValues accent2))
                target.Accent2 = accent2;
            if (values.TryGetValue("accent3", out A.ColorSchemeIndexValues accent3))
                target.Accent3 = accent3;
            if (values.TryGetValue("accent4", out A.ColorSchemeIndexValues accent4))
                target.Accent4 = accent4;
            if (values.TryGetValue("accent5", out A.ColorSchemeIndexValues accent5))
                target.Accent5 = accent5;
            if (values.TryGetValue("accent6", out A.ColorSchemeIndexValues accent6))
                target.Accent6 = accent6;
            if (values.TryGetValue("hlink", out A.ColorSchemeIndexValues hyperlink))
                target.Hyperlink = hyperlink;
            if (values.TryGetValue("folHlink", out A.ColorSchemeIndexValues followed))
                target.FollowedHyperlink = followed;
        }
    }
}
