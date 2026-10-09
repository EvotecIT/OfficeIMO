using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    /// <summary>Writes parseable optional metadata and a page window without inventing display dimensions or a thumbnail image.</summary>
    private static void WritePackageMetadata(PackagePart appPart, PackagePart customPart, PackagePart windowsPart, IReadOnlyList<VisioPage> pages) {
        WriteXml(appPart, new XElement(XName.Get("Properties", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties")));
        WriteXml(customPart, new XElement(XName.Get("Properties", "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties")));

        XNamespace ns = VisioNamespace;
        XElement windows = new(ns + "Windows");
        if (pages.Count > 0) {
            windows.Add(new XElement(ns + "Window",
                new XAttribute("ID", 0),
                new XAttribute("WindowType", "Drawing"),
                new XAttribute("WindowState", 0),
                new XAttribute("ContainerType", "Page"),
                new XAttribute("Container", pages[0].Id),
                new XAttribute("Page", pages[0].Id)));
        }
        WriteXml(windowsPart, windows);

        static void WriteXml(PackagePart part, XElement root) {
            using Stream output = part.GetStream(FileMode.Create, FileAccess.Write);
            new XDocument(root).Save(output, SaveOptions.DisableFormatting);
        }
    }
}
