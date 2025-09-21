using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VsdxPackageValidator {
        private bool IsStyle0Referenced(string? styleValue) => int.TryParse(styleValue, out var id) && id == 0;
        private XDocument? LoadXml(string path) { try { return XDocument.Load(path, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo); } catch { return null; } }
        private void SaveXml(XDocument doc, string path) {
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
            var settings = new XmlWriterSettings { Indent = true, Encoding = new UTF8Encoding(false), OmitXmlDeclaration = false };
            using var writer = XmlWriter.Create(path, settings);
            doc.Save(writer);
        }
        private bool HasDefault(XDocument doc, string ext, string contentType) => doc.Root?.Elements(nsCT + "Default").Any(e => (string?)e.Attribute("Extension") == ext && (string?)e.Attribute("ContentType") == contentType) ?? false;
        private bool HasOverride(XDocument doc, string partName, string contentType) => doc.Root?.Elements(nsCT + "Override").Any(e => (string?)e.Attribute("PartName") == partName && (string?)e.Attribute("ContentType") == contentType) ?? false;
        private void AddDefault(XDocument doc, string ext, string contentType) { if (doc.Root != null && !HasDefault(doc, ext, contentType)) doc.Root.Add(new XElement(nsCT + "Default", new XAttribute("Extension", ext), new XAttribute("ContentType", contentType))); }
        private void AddOverride(XDocument doc, string partName, string contentType) { if (doc.Root != null && !HasOverride(doc, partName, contentType)) doc.Root.Add(new XElement(nsCT + "Override", new XAttribute("PartName", partName), new XAttribute("ContentType", contentType))); }

        private void CreateDefaultContentTypes(string path) {
            var doc = new XDocument(
                new XElement(nsCT + "Types",
                    new XElement(nsCT + "Default",
                        new XAttribute("Extension", "rels"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                    new XElement(nsCT + "Default",
                        new XAttribute("Extension", "xml"),
                        new XAttribute("ContentType", "application/xml")),
                    new XElement(nsCT + "Override",
                        new XAttribute("PartName", "/visio/document.xml"),
                        new XAttribute("ContentType", CT_Document)),
                    new XElement(nsCT + "Override",
                        new XAttribute("PartName", "/visio/pages/pages.xml"),
                        new XAttribute("ContentType", CT_Pages)),
                    new XElement(nsCT + "Override",
                        new XAttribute("PartName", "/visio/pages/page1.xml"),
                        new XAttribute("ContentType", CT_Pages))));
            SaveXml(doc, path);
        }
    }
}
