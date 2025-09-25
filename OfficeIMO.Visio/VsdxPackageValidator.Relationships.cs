using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VsdxPackageValidator {
        private void ValidatePackageRelationships(string tempPath, bool fix) {
            var pkgRelsPath = Path.Combine(tempPath, "_rels", ".rels");

            if (!File.Exists(pkgRelsPath)) {
                _errors.Add("Missing /_rels/.rels");
                if (fix) {
                    CreatePackageRels(pkgRelsPath);
                    _fixes.Add("Created /_rels/.rels");
                }
                return;
            }

            var doc = LoadXml(pkgRelsPath);
            if (doc?.Root == null) {
                _errors.Add("Malformed /_rels/.rels");
                if (fix) {
                    CreatePackageRels(pkgRelsPath);
                    _fixes.Add("Recreated /_rels/.rels");
                }
                return;
            }

            var docRel = doc.Root.Elements(nsPkgRel + "Relationship")
                .FirstOrDefault(r => (string?)r.Attribute("Type") == RT_Document);

            if (docRel == null) {
                _errors.Add("No root relationship to visio/document.xml");
                if (fix) {
                    doc.Root.Add(new XElement(nsPkgRel + "Relationship",
                        new XAttribute("Id", "rIdDoc"),
                        new XAttribute("Type", RT_Document),
                        new XAttribute("Target", "visio/document.xml")));
                    SaveXml(doc, pkgRelsPath);
                    _fixes.Add("Added root -> document.xml relationship");
                }
            } else {
                string? target = (string?)docRel.Attribute("Target");
                if (!string.Equals(target, "visio/document.xml", StringComparison.OrdinalIgnoreCase)) {
                    _warnings.Add($"Root document relationship target is '{target}', expected 'visio/document.xml'");
                }
            }
        }

        private void ValidateDocumentRelationships(string tempPath, bool fix) {
            var docPath = Path.Combine(tempPath, "visio", "document.xml");
            var docRelsPath = Path.Combine(tempPath, "visio", "_rels", "document.xml.rels");

            if (!File.Exists(docPath)) {
                _errors.Add("Missing /visio/document.xml");
                if (fix) {
                    CreateMinimalDocument(docPath);
                    _fixes.Add("Created minimal /visio/document.xml");
                }
            }

            if (!File.Exists(docRelsPath)) {
                _errors.Add("Missing /visio/_rels/document.xml.rels");
                if (fix) {
                    CreateDocumentRels(docRelsPath);
                    _fixes.Add("Created /visio/_rels/document.xml.rels");
                }
                return;
            }

            var doc = LoadXml(docRelsPath);
            if (doc?.Root == null) {
                _errors.Add("Malformed /visio/_rels/document.xml.rels");
                if (fix) {
                    CreateDocumentRels(docRelsPath);
                    _fixes.Add("Recreated /visio/_rels/document.xml.rels");
                }
                return;
            }

            bool modified = false;
            var rels = doc.Root.Elements(nsPkgRel + "Relationship").ToList();

            var pagesRel = rels.FirstOrDefault(r => (string?)r.Attribute("Type") == RT_Pages);
            if (pagesRel == null) {
                _errors.Add("document.xml.rels has no pages relationship");
                if (fix) {
                    doc.Root.Add(new XElement(nsPkgRel + "Relationship",
                        new XAttribute("Id", "rIdPages"),
                        new XAttribute("Type", RT_Pages),
                        new XAttribute("Target", "pages/pages.xml")));
                    modified = true;
                    _fixes.Add("Added document -> pages relationship");
                }
            }

            var mastersRel = rels.FirstOrDefault(r => (string?)r.Attribute("Type") == RT_Masters);
            if (mastersRel != null) {
                var target = Path.Combine(tempPath, "visio",
                    ((string?)mastersRel.Attribute("Target") ?? "").Replace('/', Path.DirectorySeparatorChar));
                if (!File.Exists(target)) {
                    _errors.Add("document.xml.rels references masters but masters.xml does not exist");
                    if (fix) {
                        mastersRel.Remove();
                        modified = true;
                        _fixes.Add("Removed dangling masters relationship");
                    }
                }
            }

            if (fix && modified) {
                SaveXml(doc, docRelsPath);
            }
        }

        private void CreatePackageRels(string path) {
            var doc = new XDocument(
                new XElement(nsPkgRel + "Relationships",
                    new XElement(nsPkgRel + "Relationship",
                        new XAttribute("Id", "rIdDoc"),
                        new XAttribute("Type", RT_Document),
                        new XAttribute("Target", "visio/document.xml"))));
            SaveXml(doc, path);
        }

        private void CreateDocumentRels(string path) {
            var doc = new XDocument(
                new XElement(nsPkgRel + "Relationships",
                    new XElement(nsPkgRel + "Relationship",
                        new XAttribute("Id", "rIdPages"),
                        new XAttribute("Type", RT_Pages),
                        new XAttribute("Target", "pages/pages.xml"))));
            SaveXml(doc, path);
        }

        private void CreatePagesRels(string path) {
            var doc = new XDocument(
                new XElement(nsPkgRel + "Relationships",
                    new XElement(nsPkgRel + "Relationship",
                        new XAttribute("Id", "rId1"),
                        new XAttribute("Type", RT_Page),
                        new XAttribute("Target", "page1.xml"))));
            SaveXml(doc, path);
        }
    }
}

