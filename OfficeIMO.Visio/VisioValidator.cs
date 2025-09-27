using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Provides methods to validate Visio packages for basic compliance.
    /// </summary>
    public static class VisioValidator {
        private static readonly XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
        private static readonly XNamespace rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace pr = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";

        private const string RT_Document = "http://schemas.microsoft.com/visio/2010/relationships/document";
        private const string RT_Pages = "http://schemas.microsoft.com/visio/2010/relationships/pages";
        private const string RT_Page = "http://schemas.microsoft.com/visio/2010/relationships/page";

        private const string CT_Document = "application/vnd.ms-visio.drawing.main+xml";
        private const string CT_Pages = "application/vnd.ms-visio.pages+xml";
        private const string CT_Page = "application/vnd.ms-visio.page+xml";

        /// <summary>
        /// Validates the specified Visio package.
        /// </summary>
        /// <param name="vsdxPath">Path to the VSDX package.</param>
        /// <returns>List of validation issues.</returns>
        public static IReadOnlyList<string> Validate(string vsdxPath) {
            List<string> issues = new();
            using Package pkg = Package.Open(vsdxPath, FileMode.Open, FileAccess.Read, FileShare.Read);

            XDocument ctDoc;
            using (FileStream zipStream = File.OpenRead(vsdxPath))
            using (ZipArchive archive = new(zipStream, ZipArchiveMode.Read))
            using (Stream s = archive.GetEntry("[Content_Types].xml")!.Open()) {
                ctDoc = XDocument.Load(s);
            }
            var defaults = ctDoc.Root!.Elements(ct + "Default").ToList();
            var overrides = ctDoc.Root!.Elements(ct + "Override").ToList();

            XElement? xmlDefault = defaults.FirstOrDefault(d => (string?)d.Attribute("Extension") == "xml");
            if (xmlDefault == null || (string?)xmlDefault.Attribute("ContentType") != "application/xml") {
                issues.Add("Default for '.xml' must be 'application/xml' with per-part Overrides.");
            }

            bool HasOverride(string partName, string type) =>
                overrides.Any(o => (string?)o.Attribute("PartName") == partName && (string?)o.Attribute("ContentType") == type);

            if (!HasOverride("/visio/document.xml", CT_Document)) {
                issues.Add("Missing Override for /visio/document.xml -> application/vnd.ms-visio.drawing.main+xml.");
            }

            if (!HasOverride("/visio/pages/pages.xml", CT_Pages)) {
                issues.Add("Missing Override for /visio/pages/pages.xml -> application/vnd.ms-visio.pages+xml.");
            }

            if (!HasOverride("/visio/pages/page1.xml", CT_Page)) {
                issues.Add("Missing Override for /visio/pages/page1.xml -> application/vnd.ms-visio.page+xml.");
            }

            XDocument rootRels = GetRels(pkg, "/_rels/.rels");
            XElement? docRel = rootRels.Root!.Elements(pr + "Relationship").FirstOrDefault(r => (string?)r.Attribute("Target") == "/visio/document.xml");
            if (docRel == null || (string?)docRel.Attribute("Type") != RT_Document) {
                issues.Add("Root relationship must target /visio/document.xml with Visio document type.");
            }

            XDocument docRels = GetRels(pkg, "/visio/_rels/document.xml.rels");
            XElement? pagesRel = docRels.Root!.Elements(pr + "Relationship").FirstOrDefault(r => (string?)r.Attribute("Target") == "pages/pages.xml");
            if (pagesRel == null || (string?)pagesRel.Attribute("Type") != RT_Pages) {
                issues.Add("document.xml must relate to pages/pages.xml with visio/2010/relationships/pages.");
            }

            XDocument pagesXml = LoadXml(pkg, "/visio/pages/pages.xml");
            IReadOnlyList<XElement> pageElements = pagesXml.Root!
                .Elements(v + "Page")
                .ToList();

            if (pageElements.Count == 0) {
                issues.Add("pages.xml must contain a Page element.");
            }

            XDocument pagesRels = GetRels(pkg, "/visio/pages/_rels/pages.xml.rels");
            Dictionary<string, XElement> relationshipsById = new(StringComparer.Ordinal);
            foreach (XElement relElem in pagesRels.Root!.Elements(pr + "Relationship")) {
                string? relIdAttr = (string?)relElem.Attribute("Id");
                if (!string.IsNullOrEmpty(relIdAttr)) {
                    relationshipsById[relIdAttr] = relElem;
                }
            }

            HashSet<string> processedPageParts = new(StringComparer.OrdinalIgnoreCase);

            foreach (XElement pageElement in pageElements) {
                string pageLabel = DescribePage(pageElement);
                string? idAttr = (string?)pageElement.Attribute("ID");
                if (!int.TryParse(idAttr, out int pageId) || pageId < 0) {
                    issues.Add($"Page {pageLabel}: ID must be numeric and zero-based (e.g., 0).");
                }

                List<XElement> relChildren = pageElement.Elements(v + "Rel").ToList();
                if (relChildren.Count == 0) {
                    issues.Add($"Page {pageLabel}: Must contain <Rel r:id=\"rId#\"> child (not an attribute).");
                    continue;
                }

                foreach (XElement relChild in relChildren) {
                    string? rid = relChild.Attribute(rel + "id")?.Value;
                    if (string.IsNullOrWhiteSpace(rid) || !rid.StartsWith("rId", StringComparison.Ordinal)) {
                        issues.Add($"Page {pageLabel}: Must contain <Rel r:id=\"rId#\"> child (not an attribute).");
                        continue;
                    }

                    if (!relationshipsById.TryGetValue(rid, out XElement? relationship)) {
                        issues.Add($"Page {pageLabel}: Relationship '{rid}' not found in pages.xml.rels.");
                        continue;
                    }

                    string? type = (string?)relationship.Attribute("Type");
                    if (!string.Equals(type, RT_Page, StringComparison.Ordinal)) {
                        issues.Add($"Page {pageLabel}: Relationship '{rid}' must be of type {RT_Page}.");
                    }

                    string? target = (string?)relationship.Attribute("Target");
                    if (string.IsNullOrWhiteSpace(target)) {
                        issues.Add($"Page {pageLabel}: Relationship '{rid}' must target a page part.");
                        continue;
                    }

                    Uri partUri = PackUriHelper.ResolvePartUri(new Uri("/visio/pages/pages.xml", UriKind.Relative), new Uri(target, UriKind.Relative));
                    string partName = partUri.ToString();

                    if (!pkg.PartExists(partUri)) {
                        issues.Add($"Page {pageLabel}: Target part '{partName}' referenced by relationship '{rid}' is missing.");
                    } else if (processedPageParts.Add(partName)) {
                        XDocument pageXml = LoadXml(pkg, partName);
                        string? badId = pageXml
                            .Descendants(v + "Shape")
                            .Select(x => (string?)x.Attribute("ID"))
                            .FirstOrDefault(shapeId => !int.TryParse(shapeId, out _));
                        if (badId != null) {
                            issues.Add($"Page {pageLabel}: Shape/@ID must be numeric. Found non-numeric ID: '{badId}'.");
                        }
                    }

                    if (!HasOverride(partName, CT_Page)) {
                        issues.Add($"Page {pageLabel}: Missing Override for {partName} -> {CT_Page}.");
                    }
                }
            }

            return issues;

            static string DescribePage(XElement pageElement) {
                string id = pageElement.Attribute("ID")?.Value ?? "?";
                string? name = pageElement.Attribute("NameU")?.Value;
                name ??= pageElement.Attribute("Name")?.Value;
                return name != null ? $"ID {id} ({name})" : $"ID {id}";
            }
        }

        private static XDocument LoadXml(Package pkg, string partName) {
            using Stream s = pkg.GetPart(new Uri(partName, UriKind.Relative)).GetStream();
            return XDocument.Load(s);
        }

        private static XDocument GetRels(Package pkg, string relsPath) {
            using Stream s = pkg.GetPart(new Uri(relsPath, UriKind.Relative)).GetStream();
            return XDocument.Load(s);
        }
    }
}
