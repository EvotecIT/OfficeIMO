using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

internal static partial class VisioLegacyXmlCodec {
    private static readonly XNamespace Legacy2002 = "urn:schemas-microsoft-com:office:visio";

    private static XElement NormalizeLegacyRoot(XDocument source, VisioXmlConversionReport report) {
        XElement root = source.Root ?? throw new InvalidDataException("Visio XML has no root element.");
        if (root.Name == Legacy + "VisioDocument") return root;
        if (root.Name != Legacy2002 + "VisioDocument")
            throw new InvalidDataException("Expected a Visio 2002 or 2003 core VisioDocument root.");
        // Work on the bounded parser's private tree. Keep extension namespaces intact.
        foreach (XElement element in root.DescendantsAndSelf()) {
            if (element.Name.Namespace == Legacy2002) element.Name = Legacy + element.Name.LocalName;
        }
        report.Add("VDX_2002_INPUT", "Visio 2002 XML is normalized into the shared model; legacy XML export uses the Visio 2003 namespace.", OfficeConversionLossKind.None);
        return root;
    }

    private static void ResolveLegacyFontTable(XElement root, VisioXmlConversionReport report) {
        XElement? fonts = root.Element(Modern + "Fonts");
        if (fonts == null) return;
        XElement? faces = root.Element(Modern + "FaceNames");
        if (faces == null) { faces = new XElement(Modern + "FaceNames"); root.Add(faces); }
        var existing = faces.Elements(Modern + "FaceName").Where(face => face.Attribute("ID") != null)
            .GroupBy(face => (string)face.Attribute("ID")!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        foreach (XElement font in fonts.Elements(Modern + "FontEntry")) {
            string? id = (string?)font.Attribute("ID"), name = (string?)font.Attribute("Name");
            if (id == null || string.IsNullOrWhiteSpace(name)) continue;
            if (existing.TryGetValue(id, out XElement? face)) {
                if (!string.Equals((string?)face.Attribute("Name"), name, StringComparison.OrdinalIgnoreCase))
                    report.Add("VDX_FONT_TABLE_CONFLICT", "FontEntry and FaceName disagree for this identifier; the FaceName table controls text rendering.", OfficeConversionLossKind.Approximation, id);
                continue;
            }
            // Retain the original legacy Fonts table. The modern face is the shared loader's
            // interpretation of its name/charset, not a rewrite of the native font metadata.
            face = new XElement(Modern + "FaceName", new XAttribute("ID", id), new XAttribute("Name", name));
            if (font.Attribute("CharSet") is XAttribute charSet) face.SetAttributeValue("CharSets", charSet.Value);
            faces.Add(face); existing.Add(id, face);
        }
    }
}
