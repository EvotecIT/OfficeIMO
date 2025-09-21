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
    /// Validates and optionally fixes Visio VSDX packages for common structural issues.
    /// </summary>
    public partial class VsdxPackageValidator {
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
        private const string CT_MastersPart = "application/vnd.ms-visio.masters+xml";
        private const string CT_Master = "application/vnd.ms-visio.master+xml";

        private readonly List<string> _errors = new();
        private readonly List<string> _warnings = new();
        private readonly List<string> _fixes = new();

        /// <summary>
        /// Gets the list of validation errors.
        /// </summary>
        public IReadOnlyList<string> Errors => _errors.AsReadOnly();

        /// <summary>
        /// Gets the list of warnings encountered during validation.
        /// </summary>
        public IReadOnlyList<string> Warnings => _warnings.AsReadOnly();

        /// <summary>
        /// Gets the list of fixes applied when running in fix mode.
        /// </summary>
        public IReadOnlyList<string> Fixes => _fixes.AsReadOnly();

        /// <summary>
        /// Validates the specified VSDX file.
        /// </summary>
        /// <param name="inputPath">Path to the input VSDX file.</param>
        /// <returns><c>true</c> if no errors were found; otherwise, <c>false</c>.</returns>
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
                try { Directory.Delete(tempPath, recursive: true); } catch { /* ignore cleanup errors */ }
            }
        }

        /// <summary>
        /// Validates and fixes the specified VSDX file.
        /// </summary>
        /// <param name="inputPath">Path to the input VSDX file.</param>
        /// <param name="outputPath">Path where the fixed file will be saved.</param>
        /// <returns><c>true</c> if the file was fixed successfully; otherwise, <c>false</c>.</returns>
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
                try { Directory.Delete(tempPath, recursive: true); } catch { /* ignore cleanup errors */ }
            }
        }

        private string ExtractToTemp(string inputPath) {
            var tempBase = Path.GetTempPath();
            var rnd = Path.GetRandomFileName();
            var tempPath = Path.Combine(tempBase, $"VsdxValidator_{rnd}");
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

        /// <summary>
        /// Streaming (no-extract) validation prototype.
        /// Phase 1: content types and relationships; Phase 2: basic pages structure.
        /// </summary>
        public bool ValidateFileStreaming(string inputPath) {
            _errors.Clear();
            _warnings.Clear();
            _fixes.Clear();

            if (!File.Exists(inputPath)) {
                _errors.Add($"File not found: {inputPath}");
                return false;
            }

            using ZipArchive zip = ZipFile.OpenRead(inputPath);
            ValidateStreamingPhase1(zip);
            ValidateStreamingPhase2(zip);
            return _errors.Count == 0;
        }

        private void ValidateStreamingPhase1(ZipArchive zip) {
            // [Content_Types].xml
            var ct = LoadZipXml(zip, "[Content_Types].xml");
            if (ct?.Root == null) { _errors.Add("Missing or malformed [Content_Types].xml"); return; }

            bool hasXmlDefault = ct.Root.Elements(nsCT + "Default").Any(e => (string?)e.Attribute("Extension") == "xml" && (string?)e.Attribute("ContentType") == "application/xml");
            bool hasRelsDefault = ct.Root.Elements(nsCT + "Default").Any(e => (string?)e.Attribute("Extension") == "rels" && (string?)e.Attribute("ContentType") == "application/vnd.openxmlformats-package.relationships+xml");
            if (!hasXmlDefault) _warnings.Add("[Content_Types].xml lacks Default for xml");
            if (!hasRelsDefault) _warnings.Add("[Content_Types].xml lacks Default for rels");

            bool hasDocOverride = ct.Root.Elements(nsCT + "Override").Any(e => (string?)e.Attribute("PartName") == "/visio/document.xml" && (string?)e.Attribute("ContentType") == CT_Document);
            if (!hasDocOverride) _errors.Add("[Content_Types].xml missing override for /visio/document.xml");
            bool hasPagesOverride = ct.Root.Elements(nsCT + "Override").Any(e => (string?)e.Attribute("PartName") == "/visio/pages/pages.xml" && (string?)e.Attribute("ContentType") == CT_Pages);
            if (!hasPagesOverride) _errors.Add("[Content_Types].xml missing override for /visio/pages/pages.xml");

            // /_rels/.rels -> document.xml
            var pkgRels = LoadZipXml(zip, "_rels/.rels");
            if (pkgRels?.Root == null) { _errors.Add("Missing /_rels/.rels"); return; }
            var docRel = pkgRels.Root.Elements(nsPkgRel + "Relationship").FirstOrDefault(r => (string?)r.Attribute("Type") == RT_Document);
            if (docRel == null) { _errors.Add("No root relationship to visio/document.xml"); } else {
                string? target = (string?)docRel.Attribute("Target");
                if (!string.Equals(target, "visio/document.xml", StringComparison.OrdinalIgnoreCase)) _warnings.Add($"Root document relationship target is '{target}'");
            }

            // /visio/_rels/document.xml.rels -> pages/pages.xml
            var docRels = LoadZipXml(zip, "visio/_rels/document.xml.rels");
            if (docRels?.Root == null) { _errors.Add("Missing /visio/_rels/document.xml.rels"); return; }
            var pagesRel = docRels.Root.Elements(nsPkgRel + "Relationship").FirstOrDefault(r => (string?)r.Attribute("Type") == RT_Pages);
            if (pagesRel == null) _errors.Add("document.xml.rels has no pages relationship");

            // /visio/pages/_rels/pages.xml.rels -> at least one pageN.xml
            var pagesRels = LoadZipXml(zip, "visio/pages/_rels/pages.xml.rels");
            if (pagesRels?.Root == null) { _errors.Add("Missing /visio/pages/_rels/pages.xml.rels"); return; }
            var pageTargets = pagesRels.Root
                .Elements(nsPkgRel + "Relationship")
                .Where(r => (string?)r.Attribute("Type") == RT_Page)
                .Select(r => (string?)r.Attribute("Target") ?? string.Empty)
                .ToList();
            if (pageTargets.Count == 0) _errors.Add("pages.xml.rels has no page relationships");
            foreach (var t in pageTargets) {
                if (string.IsNullOrEmpty(t)) { _errors.Add("pages.xml.rels contains a page relationship with empty target"); continue; }
                // Normalize path relative to /visio/pages/
                string path = "visio/pages/" + t.Replace('\\', '/');
                if (zip.GetEntry(path) == null) _errors.Add($"Missing page part: /{path}");
                bool hasCt = ct.Root.Elements(nsCT + "Override").Any(e => (string?)e.Attribute("PartName") == "/" + path && (string?)e.Attribute("ContentType") == CT_Page);
                if (!hasCt) _warnings.Add($"[Content_Types].xml missing override for /{path} with CT_Page");
            }
        }

        private void ValidateStreamingPhase2(ZipArchive zip) {
            // Basic structure: /visio/document.xml and /visio/pages/pages.xml present and parseable
            if (zip.GetEntry("visio/document.xml") == null) _errors.Add("Missing /visio/document.xml");
            var pagesDoc = LoadZipXml(zip, "visio/pages/pages.xml");
            if (pagesDoc?.Root == null) { _errors.Add("Missing or malformed /visio/pages/pages.xml"); return; }

            XNamespace vNs = "http://schemas.microsoft.com/office/visio/2012/main";
            XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            var pages = pagesDoc.Root.Elements(vNs + "Page").ToList();
            if (pages.Count == 0) {
                _errors.Add("/visio/pages/pages.xml contains no Page elements");
                return;
            }

            // Load pages relationships once
            var pagesRels = LoadZipXml(zip, "visio/pages/_rels/pages.xml.rels");
            if (pagesRels?.Root == null) { _errors.Add("Missing /visio/pages/_rels/pages.xml.rels"); return; }
            var relElems = pagesRels.Root.Elements(nsPkgRel + "Relationship").ToList();
            // r:id uniqueness in pages.xml.rels
            var idGroups = relElems.Select(e => (string?)e.Attribute("Id") ?? string.Empty)
                .GroupBy(id => id, StringComparer.Ordinal)
                .Where(g => !string.IsNullOrEmpty(g.Key) && g.Count() > 1)
                .ToList();
            foreach (var g in idGroups) _errors.Add($"Duplicate relationship Id in pages.xml.rels: '{g.Key}'");

            var relsById = relElems.ToDictionary(e => (string?)e.Attribute("Id") ?? string.Empty, e => e);

            // Page ID uniqueness in pages.xml
            var idAttrGroups = pages.Select(p => (string?)p.Attribute("ID") ?? string.Empty)
                .GroupBy(id => id, StringComparer.Ordinal)
                .Where(g => !string.IsNullOrEmpty(g.Key) && g.Count() > 1)
                .ToList();
            foreach (var g in idAttrGroups) _errors.Add($"Duplicate Page @ID in pages.xml: '{g.Key}'");

            foreach (var page in pages) {
                // Page → Rel r:id resolution
                var relElem = page.Element(vNs + "Rel");
                string? rid = relElem?.Attribute(rNs + "id")?.Value;
                if (string.IsNullOrEmpty(rid)) { _errors.Add("Page element missing Rel/@r:id"); continue; }

                if (!relsById.TryGetValue(rid!, out var rel)) { _errors.Add($"pages.xml.rels missing relationship with Id={rid}"); continue; }

                string? target = (string?)rel.Attribute("Target");
                if (string.IsNullOrEmpty(target)) { _errors.Add($"Relationship Id={rid} has empty Target"); continue; }

                string pagePath = "visio/pages/" + target!.Replace('\\', '/');
                var pageEntry = zip.GetEntry(pagePath);
                if (pageEntry == null) { _errors.Add($"Missing page part: /{pagePath}"); continue; }

                // Minimal parse of pageN.xml: root and optional Shapes/Connects sections
                var pageXml = LoadZipXml(zip, pagePath);
                if (pageXml?.Root == null) { _errors.Add($"Malformed /{pagePath}"); continue; }
                if (!XName.Get(pageXml.Root.Name.LocalName, pageXml.Root.Name.NamespaceName).Equals(vNs + "PageContents")) {
                    _warnings.Add($"/{pagePath} root is '{pageXml.Root.Name}', expected 'PageContents'");
                }

                // Verify minimal PageSheet cells
                var pageSheet = page.Element(vNs + "PageSheet");
                if (pageSheet == null) {
                    _warnings.Add("Page missing PageSheet");
                } else {
                    bool hasW = pageSheet.Elements(vNs + "Cell").Any(c => (string?)c.Attribute("N") == "PageWidth");
                    bool hasH = pageSheet.Elements(vNs + "Cell").Any(c => (string?)c.Attribute("N") == "PageHeight");
                    if (!hasW) _warnings.Add("PageSheet missing PageWidth cell");
                    if (!hasH) _warnings.Add("PageSheet missing PageHeight cell");
                }

                var shapes = pageXml.Root.Element(vNs + "Shapes");
                if (shapes != null) {
                    // ensure all children are Shape nodes in the right namespace
                    foreach (var child in shapes.Elements()) {
                        if (child.Name != vNs + "Shape") {
                            _warnings.Add($"/{pagePath} contains unexpected element under Shapes: '{child.Name}'");
                            break;
                        }
                    }
                }

                var connects = pageXml.Root.Element(vNs + "Connects");
                if (connects != null) {
                    foreach (var child in connects.Elements()) {
                        if (child.Name != vNs + "Connect") {
                            _warnings.Add($"/{pagePath} contains unexpected element under Connects: '{child.Name}'");
                            break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Streaming fix: writes a new VSDX with corrected [Content_Types].xml
        /// and minimal relationships for document and pages.
        /// </summary>
        public bool FixFileStreaming(string inputPath, string outputPath) {
            _errors.Clear();
            _warnings.Clear();
            _fixes.Clear();

            if (!File.Exists(inputPath)) { _errors.Add($"File not found: {inputPath}"); return false; }
            var outDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir)) Directory.CreateDirectory(outDir);
            if (File.Exists(outputPath)) File.Delete(outputPath);

            using ZipArchive src = ZipFile.OpenRead(inputPath);
            using ZipArchive dst = ZipFile.Open(outputPath, ZipArchiveMode.Create);

            // Copy all entries except those we will recreate
            var skip = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                "[Content_Types].xml",
                "_rels/.rels",
                "visio/_rels/document.xml.rels",
                "visio/pages/_rels/pages.xml.rels",
            };
            foreach (var e in src.Entries) {
                string name = e.FullName.Replace('\\', '/');
                if (skip.Contains(name)) continue;
                var ne = dst.CreateEntry(name, CompressionLevel.Optimal);
                using var s = e.Open();
                using var t = ne.Open();
                s.CopyTo(t);
            }

            // Build and write [Content_Types].xml
            var ct = BuildOrUpdateContentTypes(src);
            SaveZipXml(dst, "[Content_Types].xml", ct);
            _fixes.Add("Updated [Content_Types].xml");

            // _rels/.rels with a single document relationship
            XNamespace relNs = nsPkgRel;
            var pkgRels = new XDocument(new XElement(relNs + "Relationships",
                new XElement(relNs + "Relationship",
                    new XAttribute("Id", "rIdDoc"),
                    new XAttribute("Type", RT_Document),
                    new XAttribute("Target", "visio/document.xml"))));
            SaveZipXml(dst, "_rels/.rels", pkgRels);
            _fixes.Add("Rebuilt /_rels/.rels");

            // visio/_rels/document.xml.rels -> pages
            var docRels = new XDocument(new XElement(relNs + "Relationships",
                new XElement(relNs + "Relationship",
                    new XAttribute("Id", "rIdPages"),
                    new XAttribute("Type", RT_Pages),
                    new XAttribute("Target", "pages/pages.xml"))));
            SaveZipXml(dst, "visio/_rels/document.xml.rels", docRels);
            _fixes.Add("Rebuilt /visio/_rels/document.xml.rels");

            // visio/pages/_rels/pages.xml.rels -> enumerate page parts
            var pages = new List<string>();
            foreach (var e in src.Entries) {
                string n = e.FullName.Replace('\\', '/');
                if (n.StartsWith("visio/pages/", StringComparison.OrdinalIgnoreCase)
                    && n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                    && !n.EndsWith("pages.xml", StringComparison.OrdinalIgnoreCase)
                    && !n.Contains("/_rels/", StringComparison.Ordinal)) {
                    pages.Add(n.Substring("visio/pages/".Length));
                }
            }
            pages.Sort(StringComparer.OrdinalIgnoreCase);

            var relsRoot = new XElement(relNs + "Relationships");
            for (int i = 0; i < pages.Count; i++) {
                relsRoot.Add(new XElement(relNs + "Relationship",
                    new XAttribute("Id", $"rId{i + 1}"),
                    new XAttribute("Type", RT_Page),
                    new XAttribute("Target", pages[i])));
            }
            var pagesRels = new XDocument(relsRoot);
            SaveZipXml(dst, "visio/pages/_rels/pages.xml.rels", pagesRels);
            _fixes.Add("Rebuilt /visio/pages/_rels/pages.xml.rels");

            // Run streaming validation on the output
            return ValidateFileStreaming(outputPath);
        }
    }
}
