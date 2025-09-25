using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VsdxPackageValidator {
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
    }
}

