using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Validates and optionally fixes Visio <c>.vsdx</c> packages.
    /// </summary>
    public class VsdxPackageValidator {
        private static readonly XNamespace nsCore = "http://schemas.microsoft.com/office/visio/2011/1/core";
        private static readonly XNamespace nsPkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace nsDocRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace nsCT = "http://schemas.openxmlformats.org/package/2006/content-types";

        private const string RT_Document = "http://schemas.microsoft.com/visio/2010/relationships/document";
        private const string RT_Pages = "http://schemas.microsoft.com/visio/2010/relationships/pages";
        private const string RT_Page = "http://schemas.microsoft.com/visio/2010/relationships/page";
        private const string RT_Masters = "http://schemas.microsoft.com/visio/2010/relationships/masters";

        private const string CT_Document = "application/vnd.ms-visio.drawing.main+xml";
        private const string CT_Pages = "application/vnd.ms-visio.pages+xml";
        private const string CT_Page = "application/vnd.ms-visio.page+xml";

        private readonly List<string> _errors = new();
        private readonly List<string> _warnings = new();
        private readonly List<string> _fixes = new();

        /// <summary>
        /// List of validation errors.
        /// </summary>
        public IReadOnlyList<string> Errors => _errors.AsReadOnly();

        /// <summary>
        /// List of validation warnings.
        /// </summary>
        public IReadOnlyList<string> Warnings => _warnings.AsReadOnly();

        /// <summary>
        /// List of fixes applied when repairing a package.
        /// </summary>
        public IReadOnlyList<string> Fixes => _fixes.AsReadOnly();

        /// <summary>
        /// Validates the specified <c>.vsdx</c> file.
        /// </summary>
        /// <param name="inputPath">Path to the file to validate.</param>
        public bool ValidateFile(string inputPath) {
            _errors.Clear();
            _warnings.Clear();
            _fixes.Clear();

            if (!File.Exists(inputPath)) {
                _errors.Add($"File not found: {inputPath}");
                return false;
            }

            var tempPath = ExtractToTemp(inputPath);
            try {
                ValidatePackageStructure(tempPath);
                return _errors.Count == 0;
            } finally {
                Directory.Delete(tempPath, recursive: true);
            }
        }

        /// <summary>
        /// Validates and fixes the specified file, writing the result to <paramref name="outputPath"/>.
        /// </summary>
        /// <param name="inputPath">Path to the file to fix.</param>
        /// <param name="outputPath">Path to save the fixed file.</param>
        public bool FixFile(string inputPath, string outputPath) {
            _errors.Clear();
            _warnings.Clear();
            _fixes.Clear();

            if (!File.Exists(inputPath)) {
                _errors.Add($"File not found: {inputPath}");
                return false;
            }

            var tempPath = ExtractToTemp(inputPath);
            try {
                ValidateAndFix(tempPath);
                
                if (File.Exists(outputPath)) {
                    File.Delete(outputPath);
                }
                
                ZipFile.CreateFromDirectory(tempPath, outputPath, CompressionLevel.Optimal, includeBaseDirectory: false);
                return true;
            } finally {
                Directory.Delete(tempPath, recursive: true);
            }
        }

        private string ExtractToTemp(string inputPath) {
            var tempPath = Path.Combine(Path.GetTempPath(), "VsdxValidator_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempPath);
            ZipFile.ExtractToDirectory(inputPath, tempPath);
            return tempPath;
        }

        private void ValidatePackageStructure(string tempPath) {
            ValidateContentTypes(tempPath, fix: false);
            ValidatePackageRelationships(tempPath, fix: false);
            ValidateDocumentRelationships(tempPath, fix: false);
            ValidatePagesStructure(tempPath, fix: false);
            ValidateStyleReferences(tempPath, fix: false);
        }

        private void ValidateAndFix(string tempPath) {
            ValidateContentTypes(tempPath, fix: true);
            ValidatePackageRelationships(tempPath, fix: true);
            ValidateDocumentRelationships(tempPath, fix: true);
            ValidatePagesStructure(tempPath, fix: true);
            ValidateStyleReferences(tempPath, fix: true);
        }

        private void ValidateContentTypes(string tempPath, bool fix) {
            var contentTypesPath = Path.Combine(tempPath, "[Content_Types].xml");
            
            if (!File.Exists(contentTypesPath)) {
                _errors.Add("Missing [Content_Types].xml");
                if (fix) {
                    CreateDefaultContentTypes(contentTypesPath);
                    _fixes.Add("Created [Content_Types].xml with default content");
                }
                return;
            }

            var doc = LoadXml(contentTypesPath);
            if (doc?.Root == null) {
                _errors.Add("Malformed [Content_Types].xml");
                if (fix) {
                    CreateDefaultContentTypes(contentTypesPath);
                    _fixes.Add("Recreated [Content_Types].xml");
                }
                return;
            }

            bool modified = false;
            
            if (!HasDefault(doc, "rels", "application/vnd.openxmlformats-package.relationships+xml")) {
                _errors.Add("Missing default content type for .rels");
                if (fix) {
                    AddDefault(doc, "rels", "application/vnd.openxmlformats-package.relationships+xml");
                    modified = true;
                    _fixes.Add("Added default content type for .rels");
                }
            }

            if (!HasDefault(doc, "xml", "application/xml")) {
                _warnings.Add("Missing default content type for .xml");
                if (fix) {
                    AddDefault(doc, "xml", "application/xml");
                    modified = true;
                    _fixes.Add("Added default content type for .xml");
                }
            }

            if (!HasOverride(doc, "/visio/document.xml", CT_Document)) {
                _errors.Add("Missing override for /visio/document.xml");
                if (fix) {
                    AddOverride(doc, "/visio/document.xml", CT_Document);
                    modified = true;
                    _fixes.Add("Added override for /visio/document.xml");
                }
            }

            if (File.Exists(Path.Combine(tempPath, "visio", "pages", "pages.xml"))) {
                if (!HasOverride(doc, "/visio/pages/pages.xml", CT_Pages)) {
                    _errors.Add("Missing override for /visio/pages/pages.xml");
                    if (fix) {
                        AddOverride(doc, "/visio/pages/pages.xml", CT_Pages);
                        modified = true;
                        _fixes.Add("Added override for /visio/pages/pages.xml");
                    }
                }
            }

            var pagesDir = Path.Combine(tempPath, "visio", "pages");
            if (Directory.Exists(pagesDir)) {
                foreach (var pagePath in Directory.GetFiles(pagesDir, "page*.xml")) {
                    var fileName = Path.GetFileName(pagePath);
                    var partName = $"/visio/pages/{fileName}";
                    if (!HasOverride(doc, partName, CT_Page)) {
                        _errors.Add($"Missing override for {partName}");
                        if (fix) {
                            AddOverride(doc, partName, CT_Page);
                            modified = true;
                            _fixes.Add($"Added override for {partName}");
                        }
                    }
                }
            }

            if (fix && modified) {
                SaveXml(doc, contentTypesPath);
            }
        }

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
                .FirstOrDefault(r => (string)r.Attribute("Type") == RT_Document);

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
                var target = (string)docRel.Attribute("Target");
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
            
            var pagesRel = rels.FirstOrDefault(r => 
                (string)r.Attribute("Type") == RT_Pages);
            
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

            var mastersRel = rels.FirstOrDefault(r => 
                (string)r.Attribute("Type") == RT_Masters);
            
            if (mastersRel != null) {
                var target = Path.Combine(tempPath, "visio", 
                    ((string)mastersRel.Attribute("Target") ?? "").Replace('/', Path.DirectorySeparatorChar));
                
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

        private void ValidatePagesStructure(string tempPath, bool fix) {
            var pagesXmlPath = Path.Combine(tempPath, "visio", "pages", "pages.xml");
            var pagesRelsPath = Path.Combine(tempPath, "visio", "pages", "_rels", "pages.xml.rels");

            if (!File.Exists(pagesXmlPath)) {
                _errors.Add("Missing /visio/pages/pages.xml");
                if (fix) {
                    CreateMinimalPages(pagesXmlPath);
                    CreatePagesRels(pagesRelsPath);
                    CreateMinimalPage1(Path.Combine(tempPath, "visio", "pages", "page1.xml"));
                    _fixes.Add("Created pages.xml, pages.xml.rels, and page1.xml");
                }
                return;
            }

            if (!File.Exists(pagesRelsPath)) {
                _errors.Add("Missing /visio/pages/_rels/pages.xml.rels");
                if (fix) {
                    CreatePagesRels(pagesRelsPath);
                    _fixes.Add("Created /visio/pages/_rels/pages.xml.rels");
                }
            }

            var pagesDoc = LoadXml(pagesXmlPath);
            var pagesRelsDoc = File.Exists(pagesRelsPath) ? LoadXml(pagesRelsPath) : null;

            if (pagesDoc?.Root != null && pagesRelsDoc?.Root != null) {
                var relElements = pagesDoc.Descendants(nsCore + "Rel").ToList();
                var relMap = pagesRelsDoc.Root.Elements(nsPkgRel + "Relationship")
                    .ToDictionary(e => (string)e.Attribute("Id") ?? "", 
                                  e => (string)e.Attribute("Target") ?? "");

                int pageIndex = 1;
                bool modified = false;
                
                foreach (var rel in relElements) {
                    var id = (string)rel.Attribute(nsDocRel + "id") ?? "";
                    if (string.IsNullOrWhiteSpace(id)) {
                        _errors.Add("A <Rel> in pages.xml has no r:id");
                        continue;
                    }

                    if (!relMap.TryGetValue(id, out var target) || string.IsNullOrWhiteSpace(target)) {
                        _errors.Add($"pages.xml references r:id '{id}' but no mapping exists");
                        if (fix) {
                            var targetName = $"page{pageIndex}.xml";
                            pagesRelsDoc.Root.Add(new XElement(nsPkgRel + "Relationship",
                                new XAttribute("Id", id),
                                new XAttribute("Type", RT_Page),
                                new XAttribute("Target", targetName)));
                            
                            var pagePath = Path.Combine(tempPath, "visio", "pages", targetName);
                            if (!File.Exists(pagePath)) {
                                CreateMinimalPage1(pagePath);
                            }
                            
                            modified = true;
                            _fixes.Add($"Added r:id '{id}' -> '{targetName}' and ensured page exists");
                        }
                    }
                    pageIndex++;
                }

                if (fix && modified) {
                    SaveXml(pagesRelsDoc, pagesRelsPath);
                }
            }
        }

        private void ValidateStyleReferences(string tempPath, bool fix) {
            var docPath = Path.Combine(tempPath, "visio", "document.xml");
            var pagesPath = Path.Combine(tempPath, "visio", "pages", "pages.xml");

            if (!File.Exists(docPath) || !File.Exists(pagesPath)) {
                return;
            }

            var docDoc = LoadXml(docPath);
            var pagesDoc = LoadXml(pagesPath);

            if (docDoc?.Root == null || pagesDoc?.Root == null) {
                return;
            }

            var styleSheets = docDoc.Root.Element(nsCore + "StyleSheets");
            bool hasStyle0 = styleSheets?.Elements(nsCore + "StyleSheet")
                .Any(s => (string)s.Attribute("ID") == "0") ?? false;

            var pageSheet = pagesDoc.Descendants(nsCore + "PageSheet").FirstOrDefault();
            if (pageSheet != null) {
                var lineStyle = (string)pageSheet.Attribute("LineStyle");
                var fillStyle = (string)pageSheet.Attribute("FillStyle");
                var textStyle = (string)pageSheet.Attribute("TextStyle");

                bool needsStyle0 = IsStyle0Referenced(lineStyle) || 
                                   IsStyle0Referenced(fillStyle) || 
                                   IsStyle0Referenced(textStyle);

                if (needsStyle0 && !hasStyle0) {
                    _warnings.Add("PageSheet references style ID 0 but no <StyleSheet ID=\"0\"> exists");
                    if (fix) {
                        if (styleSheets == null) {
                            styleSheets = new XElement(nsCore + "StyleSheets");
                            docDoc.Root.AddFirst(styleSheets);
                        }

                        var style0 = new XElement(nsCore + "StyleSheet",
                            new XAttribute("ID", "0"),
                            new XAttribute("NameU", "No Style"),
                            new XAttribute("Name", "No Style"),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "EnableLineProps"), 
                                new XAttribute("V", "1")),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "EnableFillProps"), 
                                new XAttribute("V", "1")),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "EnableTextProps"), 
                                new XAttribute("V", "1")));
                        
                        styleSheets.AddFirst(style0);
                        SaveXml(docDoc, docPath);
                        _fixes.Add("Added minimal <StyleSheet ID=\"0\">");
                    }
                }
            }
        }

        private bool IsStyle0Referenced(string styleValue) {
            return !string.IsNullOrEmpty(styleValue) && 
                   int.TryParse(styleValue, out var id) && 
                   id == 0;
        }

        private XDocument LoadXml(string path) {
            try {
                return XDocument.Load(path, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
            } catch {
                return null;
            }
        }

        private void SaveXml(XDocument doc, string path) {
            var dir = Path.GetDirectoryName(path);
            if (!Directory.Exists(dir)) {
                Directory.CreateDirectory(dir);
            }

            var settings = new XmlWriterSettings {
                Indent = true,
                Encoding = new UTF8Encoding(false),
                OmitXmlDeclaration = false
            };

            using (var writer = XmlWriter.Create(path, settings)) {
                doc.Save(writer);
            }
        }

        private bool HasDefault(XDocument doc, string ext, string contentType) {
            return doc.Root.Elements(nsCT + "Default").Any(e =>
                (string)e.Attribute("Extension") == ext &&
                (string)e.Attribute("ContentType") == contentType);
        }

        private bool HasOverride(XDocument doc, string partName, string contentType) {
            return doc.Root.Elements(nsCT + "Override").Any(e =>
                (string)e.Attribute("PartName") == partName &&
                (string)e.Attribute("ContentType") == contentType);
        }

        private void AddDefault(XDocument doc, string ext, string contentType) {
            if (!HasDefault(doc, ext, contentType)) {
                doc.Root.Add(new XElement(nsCT + "Default",
                    new XAttribute("Extension", ext),
                    new XAttribute("ContentType", contentType)));
            }
        }

        private void AddOverride(XDocument doc, string partName, string contentType) {
            if (!HasOverride(doc, partName, contentType)) {
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

        private void CreateMinimalDocument(string path) {
            var doc = new XDocument(
                new XElement(nsCore + "VisioDocument",
                    new XAttribute(XNamespace.Xmlns + "r", nsDocRel.NamespaceName),
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    new XElement(nsCore + "StyleSheets",
                        new XElement(nsCore + "StyleSheet",
                            new XAttribute("ID", "0"),
                            new XAttribute("NameU", "No Style"),
                            new XAttribute("Name", "No Style"),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "EnableLineProps"), 
                                new XAttribute("V", "1")),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "EnableFillProps"), 
                                new XAttribute("V", "1")),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "EnableTextProps"), 
                                new XAttribute("V", "1"))))));
            SaveXml(doc, path);
        }

        private void CreateMinimalPages(string path) {
            var doc = new XDocument(
                new XElement(nsCore + "Pages",
                    new XAttribute(XNamespace.Xmlns + "r", nsDocRel.NamespaceName),
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    new XElement(nsCore + "Page",
                        new XAttribute("ID", "0"),
                        new XAttribute("NameU", "Page-1"),
                        new XAttribute("Name", "Page-1"),
                        new XElement(nsCore + "PageSheet",
                            new XAttribute("LineStyle", "0"),
                            new XAttribute("FillStyle", "0"),
                            new XAttribute("TextStyle", "0"),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "PageWidth"), 
                                new XAttribute("V", "8.5")),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "PageHeight"), 
                                new XAttribute("V", "11")),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "PageScale"), 
                                new XAttribute("V", "1")),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "DrawingScale"), 
                                new XAttribute("V", "1"))),
                        new XElement(nsCore + "Rel", 
                            new XAttribute(nsDocRel + "id", "rId1")))));
            SaveXml(doc, path);
        }

        private void CreateMinimalPage1(string path) {
            var doc = new XDocument(
                new XElement(nsCore + "PageContents",
                    new XAttribute(XNamespace.Xmlns + "r", nsDocRel.NamespaceName),
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    new XElement(nsCore + "Shapes",
                        new XElement(nsCore + "Shape",
                            new XAttribute("ID", "1"),
                            new XAttribute("Type", "Shape"),
                            new XAttribute("LineStyle", "0"),
                            new XAttribute("FillStyle", "0"),
                            new XAttribute("TextStyle", "0"),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "PinX"), 
                                new XAttribute("V", "4")),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "PinY"), 
                                new XAttribute("V", "5.5")),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "Width"), 
                                new XAttribute("V", "2")),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "Height"), 
                                new XAttribute("V", "1")),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "LocPinX"), 
                                new XAttribute("V", "1")),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "LocPinY"), 
                                new XAttribute("V", "0.5")),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "Angle"), 
                                new XAttribute("V", "0")),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "FillPattern"), 
                                new XAttribute("V", "1")),
                            new XElement(nsCore + "Cell", 
                                new XAttribute("N", "LinePattern"), 
                                new XAttribute("V", "1")),
                            new XElement(nsCore + "Section", 
                                new XAttribute("N", "Geometry"), 
                                new XAttribute("IX", "0"),
                                new XElement(nsCore + "Row", 
                                    new XAttribute("T", "RelMoveTo"), 
                                    new XAttribute("IX", "1"),
                                    new XElement(nsCore + "Cell", 
                                        new XAttribute("N", "X"), 
                                        new XAttribute("V", "0")),
                                    new XElement(nsCore + "Cell", 
                                        new XAttribute("N", "Y"), 
                                        new XAttribute("V", "0"))),
                                new XElement(nsCore + "Row", 
                                    new XAttribute("T", "RelLineTo"), 
                                    new XAttribute("IX", "2"),
                                    new XElement(nsCore + "Cell", 
                                        new XAttribute("N", "X"), 
                                        new XAttribute("V", "1")),
                                    new XElement(nsCore + "Cell", 
                                        new XAttribute("N", "Y"), 
                                        new XAttribute("V", "0"))),
                                new XElement(nsCore + "Row", 
                                    new XAttribute("T", "RelLineTo"), 
                                    new XAttribute("IX", "3"),
                                    new XElement(nsCore + "Cell", 
                                        new XAttribute("N", "X"), 
                                        new XAttribute("V", "1")),
                                    new XElement(nsCore + "Cell", 
                                        new XAttribute("N", "Y"), 
                                        new XAttribute("V", "1"))),
                                new XElement(nsCore + "Row", 
                                    new XAttribute("T", "RelLineTo"), 
                                    new XAttribute("IX", "4"),
                                    new XElement(nsCore + "Cell", 
                                        new XAttribute("N", "X"), 
                                        new XAttribute("V", "0")),
                                    new XElement(nsCore + "Cell", 
                                        new XAttribute("N", "Y"), 
                                        new XAttribute("V", "1"))),
                                new XElement(nsCore + "Row", 
                                    new XAttribute("T", "RelLineTo"), 
                                    new XAttribute("IX", "5"),
                                    new XElement(nsCore + "Cell", 
                                        new XAttribute("N", "X"), 
                                        new XAttribute("V", "0")),
                                    new XElement(nsCore + "Cell", 
                                        new XAttribute("N", "Y"), 
                                        new XAttribute("V", "0")))),
                            new XElement(nsCore + "Text", "Sample text")))));
            SaveXml(doc, path);
        }
    }
}