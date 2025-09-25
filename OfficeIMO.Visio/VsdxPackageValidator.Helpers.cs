using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VsdxPackageValidator {
        private bool IsStyle0Referenced(string? styleValue) => int.TryParse(styleValue, out var id) && id == 0;

        private static readonly XmlReaderSettings SecureXmlReaderSettings = new() {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = 10_000_000,
            MaxCharactersFromEntities = 0,
        };

        private XDocument? LoadXml(string path) {
            try {
                using FileStream fs = File.OpenRead(path);
                using XmlReader xr = XmlReader.Create(fs, SecureXmlReaderSettings);
                return XDocument.Load(xr, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
            } catch {
                return null;
            }
        }

        private XDocument? LoadZipXml(ZipArchive zip, string entryPath) {
            try {
                ZipArchiveEntry? entry = zip.GetEntry(entryPath);
                if (entry == null) return null;
                using Stream s = entry.Open();
                using XmlReader xr = XmlReader.Create(s, SecureXmlReaderSettings);
                return XDocument.Load(xr, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
            } catch {
                return null;
            }
        }

        private void SaveXml(XDocument doc, string path) {
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
            var settings = new XmlWriterSettings { Indent = true, Encoding = new UTF8Encoding(false), OmitXmlDeclaration = false };
            using var writer = XmlWriter.Create(path, settings);
            doc.Save(writer);
        }

        private void SaveZipXml(ZipArchive zip, string entryPath, XDocument doc) {
            // Normalize path
            string norm = entryPath.Replace('\\', '/');
            // Remove if exists
            zip.GetEntry(norm)?.Delete();
            var entry = zip.CreateEntry(norm, CompressionLevel.Optimal);
            var settings = new XmlWriterSettings { Indent = false, Encoding = new UTF8Encoding(false), OmitXmlDeclaration = false };
            using var s = entry.Open();
            using var w = XmlWriter.Create(s, settings);
            doc.Save(w);
        }

        private IEnumerable<string> ScanPageParts(ZipArchive zip) {
            foreach (var e in zip.Entries) {
                string n = e.FullName.Replace('\\', '/');
                if (n.StartsWith("visio/pages/", StringComparison.OrdinalIgnoreCase)
                    && n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                    && !n.EndsWith("pages.xml", StringComparison.OrdinalIgnoreCase)
                    && n.IndexOf("/_rels/", StringComparison.Ordinal) < 0) {
                    yield return n;
                }
            }
        }

        private XDocument BuildOrUpdateContentTypes(ZipArchive zip) {
            var ct = LoadZipXml(zip, "[Content_Types].xml");
            if (ct?.Root == null) {
                ct = new XDocument(new XElement(nsCT + "Types"));
            }

            // Ensure defaults
            AddDefault(ct, "rels", "application/vnd.openxmlformats-package.relationships+xml");
            AddDefault(ct, "xml", "application/xml");

            // Ensure overrides for document and pages.xml
            AddOverride(ct, "/visio/document.xml", CT_Document);
            AddOverride(ct, "/visio/pages/pages.xml", CT_Pages);

            // Pages: from pages.xml.rels if present; else scan entries
            var pagesRels = LoadZipXml(zip, "visio/pages/_rels/pages.xml.rels");
            List<string> pageParts = new();
            if (pagesRels?.Root != null) {
                foreach (var r in pagesRels.Root.Elements(nsPkgRel + "Relationship").Where(r => (string?)r.Attribute("Type") == RT_Page)) {
                    string? t = (string?)r.Attribute("Target");
                    if (!string.IsNullOrEmpty(t)) {
                        string path = "/visio/pages/" + t!.Replace('\\', '/');
                        pageParts.Add(path);
                    }
                }
            }
            if (pageParts.Count == 0) {
                pageParts.AddRange(ScanPageParts(zip).Select(p => "/" + p));
            }
            foreach (var p in pageParts.Distinct(StringComparer.OrdinalIgnoreCase)) {
                AddOverride(ct, p, CT_Page);
            }

            // Masters (if present)
            if (zip.GetEntry("visio/masters/masters.xml") != null) {
                AddOverride(ct, "/visio/masters/masters.xml", "application/vnd.ms-visio.masters+xml");
            }
            foreach (var e in zip.Entries) {
                string n = e.FullName.Replace('\\', '/');
                if (n.StartsWith("visio/masters/", StringComparison.OrdinalIgnoreCase)
                    && n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                    && n != "visio/masters/masters.xml"
                    && n.IndexOf("/_rels/", StringComparison.Ordinal) < 0) {
                    AddOverride(ct, "/" + n, "application/vnd.ms-visio.master+xml");
                }
            }

            return ct;
        }
        private bool HasDefault(XDocument doc, string ext, string contentType) {
            return doc.Root?
                .Elements(nsCT + "Default")
                .Any(e => (string?)e.Attribute("Extension") == ext && (string?)e.Attribute("ContentType") == contentType)
                ?? false;
        }

        private bool HasOverride(XDocument doc, string partName, string contentType) {
            return doc.Root?
                .Elements(nsCT + "Override")
                .Any(e => (string?)e.Attribute("PartName") == partName && (string?)e.Attribute("ContentType") == contentType)
                ?? false;
        }

        private void AddDefault(XDocument doc, string ext, string contentType) {
            if (doc.Root != null && !HasDefault(doc, ext, contentType)) {
                doc.Root.Add(new XElement(nsCT + "Default",
                    new XAttribute("Extension", ext),
                    new XAttribute("ContentType", contentType)));
            }
        }

        private void AddOverride(XDocument doc, string partName, string contentType) {
            if (doc.Root != null && !HasOverride(doc, partName, contentType)) {
                doc.Root.Add(new XElement(nsCT + "Override",
                    new XAttribute("PartName", partName),
                    new XAttribute("ContentType", contentType)));
            }
        }

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
                        new XAttribute("ContentType", CT_Page))));
            SaveXml(doc, path);
        }
    }
}
