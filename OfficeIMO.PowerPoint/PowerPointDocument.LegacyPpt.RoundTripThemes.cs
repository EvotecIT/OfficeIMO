using OfficeIMO.Core.Internal.OpenXml;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        /// <summary>
        /// Gives a slide master, notes master or handout master the binary round-trip theme: the theme part's content and
        /// the colour map of the master root.
        /// </summary>
        private static void ApplyLegacyMasterRoundTripTheme(LegacyPptImportParts parts,
            PowerPointPartRef target, LegacyPptRoundTripTheme? source) {
            if (source == null) return;
            if (!source.IsOverride && !string.IsNullOrWhiteSpace(source.ThemeXml)) {
                string xml = NormalizeLegacyThemeRootXml(source.ThemeXml!);
                PowerPointPartRef themePart = target.RelatedPart(LegacyThemeRelationshipType);
                if (themePart.IsNone) {
                    parts.AddXmlPart(target, LegacyThemeRelationshipType, LegacyPartFolder(target) + "theme/theme.xml",
                        LegacyThemeContentType, xml, out _);
                } else {
                    OfficeOpcEditablePackage package = parts.Package;
                    byte[] content = Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"utf-8\"?>" + xml);
                    package.ReplacePart(themePart.Name, content, content.Length);
                    themePart.MarkEdited();
                }
            }
            OfficeXmlNode root = target.Root;
            OfficeXmlChildOrder? order = root.LocalName switch {
                "sldMaster" => LegacyPptImportXml.SlideMasterOrder,
                "notesMaster" => LegacyPptImportXml.NotesMasterOrder,
                "handoutMaster" => LegacyPptImportXml.HandoutMasterOrder,
                _ => null
            };
            if (order == null || !TryReadLegacyColorMap(source.ColorMappingXml,
                    out IReadOnlyDictionary<string, string> values)) return;
            OfficeXmlNode map = root.GetOrAddChild(LegacyPptXml.P, "clrMap", order);
            ApplyLegacyColorMapValues(map, values);
            target.MarkEdited();
        }

        /// <summary>
        /// Gives a slide, layout or notes slide the binary round-trip theme as a theme override part, and its colour map
        /// override.
        /// </summary>
        private static void ApplyLegacyRoundTripTheme(LegacyPptImportParts parts,
            PowerPointPartRef target, LegacyPptRoundTripTheme? source) {
            if (source == null) return;
            OfficeXmlNode root = target.Root;
            OfficeXmlChildOrder? order = root.LocalName switch {
                "sld" => LegacyPptImportXml.SlideOrder,
                "sldLayout" => LegacyPptImportXml.SlideLayoutOrder,
                "notes" => LegacyPptImportXml.NotesSlideOrder,
                _ => null
            };
            if (order == null) {
                throw new ArgumentException(
                    "The target part cannot own a theme override.",
                    nameof(target));
            }
            if (!string.IsNullOrWhiteSpace(source.ThemeXml)) {
                PowerPointPartRef overridePart = target.RelatedPart(LegacyThemeOverrideRelationshipType);
                if (overridePart.IsNone) {
                    parts.AddXmlPart(target, LegacyThemeOverrideRelationshipType,
                        LegacyPartFolder(target) + "theme/themeOverride.xml", LegacyThemeOverrideContentType,
                        LegacyPptXml.A, "themeOverride", themeOverride => WriteLegacyThemeOverride(themeOverride, source), out _);
                } else {
                    OfficeXmlNode themeOverride = overridePart.Root;
                    foreach (OfficeXmlNode child in themeOverride.Children().ToList()) child.Remove();
                    WriteLegacyThemeOverride(themeOverride, source);
                    overridePart.MarkEdited();
                }
            }

            if (!TryCreateLegacyColorMapOverride(source.ColorMappingXml, out XElement? mapping)) return;
            OfficeXmlNode colorMap = root.SetChild(LegacyPptXml.P, "clrMapOvr", order);
            if (mapping == null) {
                colorMap.AddA("masterClrMapping");
            } else {
                OfficeXmlNode overrideMapping = colorMap.AddA("overrideClrMapping");
                foreach (XAttribute attribute in mapping.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration)) {
                    overrideMapping.Set(attribute.Name.LocalName, attribute.Value);
                }
            }
            target.MarkEdited();
        }

        // The colour, font and format schemes of a theme, or the content of a theme override, under a:themeOverride.
        private static void WriteLegacyThemeOverride(OfficeXmlNode target,
            LegacyPptRoundTripTheme source) {
            OfficeXmlNode parsed = LegacyPptImportXml.Parse(NormalizeLegacyThemeRootXml(source.ThemeXml!));
            if (source.IsOverride) {
                foreach (OfficeXmlNode child in parsed.Children().ToList()) target.Append(child);
                return;
            }
            OfficeXmlNode elements = parsed.Child(LegacyPptXml.A, "themeElements");
            foreach (string name in new[] { "clrScheme", "fontScheme", "fmtScheme" }) {
                OfficeXmlNode scheme = elements.Child(LegacyPptXml.A, name);
                if (!scheme.IsNone) target.Append(scheme);
            }
        }

        private static bool TryReadLegacyColorMap(string? xml,
            out IReadOnlyDictionary<string, string> values) {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            values = result;
            if (string.IsNullOrWhiteSpace(xml)) return false;
            XElement root = XElement.Parse(xml, LoadOptions.None);
            if (root.Name.LocalName != "clrMap") return false;
            foreach (XAttribute attribute in root.Attributes().Where(
                         attribute => !attribute.IsNamespaceDeclaration)) {
                if (IsLegacyColorSchemeIndex(attribute.Value)) {
                    result[attribute.Name.LocalName] = attribute.Value;
                }
            }
            return result.Count > 0;
        }

        private static string NormalizeLegacyThemeRootXml(string xml) =>
            XElement.Parse(xml, LoadOptions.PreserveWhitespace)
                .ToString(SaveOptions.DisableFormatting);

        // The mapping of a colour map override: null for the master mapping, the override mapping otherwise; false when
        // the markup has no mapping.
        private static bool TryCreateLegacyColorMapOverride(
            string? xml, out XElement? mapping) {
            mapping = null;
            if (string.IsNullOrWhiteSpace(xml)) return false;
            XElement root = XElement.Parse(xml, LoadOptions.None);
            XElement? found = root.Name.LocalName == "clrMap"
                ? root
                : root.Elements().FirstOrDefault(element =>
                    element.Name.LocalName is "masterClrMapping"
                        or "overrideClrMapping");
            if (found == null) return false;
            if (found.Name.LocalName == "masterClrMapping") return true;
            mapping = found;
            return true;
        }

        private static bool IsLegacyColorSchemeIndex(string value) => value is "dk1" or "lt1" or "dk2" or "lt2"
            or "accent1" or "accent2" or "accent3" or "accent4" or "accent5" or "accent6" or "hlink" or "folHlink";

        private static readonly string[] LegacyColorMapAttributes = {
            "bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"
        };

        private static void ApplyLegacyColorMapValues(OfficeXmlNode target,
            IReadOnlyDictionary<string, string> values) {
            foreach (string name in LegacyColorMapAttributes) {
                if (values.TryGetValue(name, out string? value)) target.Set(name, value);
            }
        }
    }
}
