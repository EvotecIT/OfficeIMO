using System.Globalization;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

/// <summary>Creates native generated geometry sections without changing imported row identities.</summary>
internal static class VisioGeometrySection {
    internal static XElement CreateGenerated(XNamespace ns, bool noFill = false) => new(ns + "Section",
        new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
        CreateCell(ns, "NoFill", noFill ? "1" : "0"), CreateCell(ns, "NoLine", "0"),
        CreateCell(ns, "NoShow", "0"), CreateCell(ns, "NoSnap", "0"), CreateCell(ns, "NoQuickDrag", "0"));

    internal static void WriteGeneratedStart(XmlWriter writer, string ns) {
        XElement section = CreateGenerated(ns);
        writer.WriteStartElement("Section", ns);
        foreach (XAttribute attribute in section.Attributes()) writer.WriteAttributeString(attribute.Name.LocalName, attribute.Value);
        foreach (XElement cell in section.Elements()) cell.WriteTo(writer);
    }

    internal static void WriteRowStart(XmlWriter writer, string ns, string type, int index) {
        writer.WriteStartElement("Row", ns);
        writer.WriteAttributeString("T", type);
        writer.WriteAttributeString("IX", index.ToString(CultureInfo.InvariantCulture));
    }

    private static XElement CreateCell(XNamespace ns, string name, string value) =>
        new(ns + "Cell", new XAttribute("N", name), new XAttribute("V", value));
}
