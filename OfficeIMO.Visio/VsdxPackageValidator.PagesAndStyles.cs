using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VsdxPackageValidator {
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
                    .ToDictionary(e => (string?)e.Attribute("Id") ?? "",
                                  e => (string?)e.Attribute("Target") ?? "");

                int pageIndex = 1;
                bool modified = false;

                foreach (var rel in relElements) {
                    var id = (string?)rel.Attribute(nsDocRel + "id") ?? "";
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
                .Any(s => (string?)s.Attribute("ID") == "0") ?? false;

            var pageSheet = pagesDoc.Descendants(nsCore + "PageSheet").FirstOrDefault();
            if (pageSheet != null) {
                string? lineStyle = (string?)pageSheet.Attribute("LineStyle");
                string? fillStyle = (string?)pageSheet.Attribute("FillStyle");
                string? textStyle = (string?)pageSheet.Attribute("TextStyle");

                bool needsStyle0 = IsStyle0Referenced(lineStyle) || IsStyle0Referenced(fillStyle) || IsStyle0Referenced(textStyle);

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
                            new XElement(nsCore + "Cell", new XAttribute("N", "EnableLineProps"), new XAttribute("V", "1")),
                            new XElement(nsCore + "Cell", new XAttribute("N", "EnableFillProps"), new XAttribute("V", "1")),
                            new XElement(nsCore + "Cell", new XAttribute("N", "EnableTextProps"), new XAttribute("V", "1")));

                        styleSheets.AddFirst(style0);
                        SaveXml(docDoc, docPath);
                        _fixes.Add("Added minimal <StyleSheet ID=\"0\">");
                    }
                }
            }
        }

        private void CreateMinimalDocument(string path) {
            var doc = new XDocument(
                new XElement(nsCore + "VisioDocument",
                    new XElement(nsCore + "DocumentSettings",
                        new XAttribute("TopPage", 0),
                        new XAttribute("DefaultTextStyle", 0),
                        new XAttribute("DefaultLineStyle", 0),
                        new XAttribute("DefaultFillStyle", 0),
                        new XAttribute("DefaultGuideStyle", 4),
                        new XElement(nsCore + "GlueSettings", 9),
                        new XElement(nsCore + "SnapSettings", 295),
                        new XElement(nsCore + "SnapExtensions", 34),
                        new XElement(nsCore + "SnapAngles"),
                        new XElement(nsCore + "DynamicGridEnabled", 1),
                        new XElement(nsCore + "ProtectStyles", 0),
                        new XElement(nsCore + "ProtectShapes", 0),
                        new XElement(nsCore + "ProtectMasters", 0),
                        new XElement(nsCore + "ProtectBkgnds", 0))));
            SaveXml(doc, path);
        }

        private void CreateMinimalPages(string path) {
            var doc = new XDocument(
                new XElement(nsCore + "Pages",
                    new XElement(nsCore + "Page",
                        new XAttribute("ID", 0),
                        new XAttribute("NameU", "Page-1"),
                        new XAttribute("Name", "Page-1"),
                        new XAttribute("ViewScale", -1),
                        new XAttribute("ViewCenterX", 4.12),
                        new XAttribute("ViewCenterY", 5.85),
                        new XElement(nsCore + "PageSheet",
                            new XAttribute("LineStyle", 0),
                            new XAttribute("FillStyle", 0),
                            new XAttribute("TextStyle", 0)),
                        new XElement(nsCore + "Rel",
                            new XAttribute(nsPkgRel + "id", "rId1")))));
            SaveXml(doc, path);
        }

        private void CreateMinimalPage1(string path) {
            var doc = new XDocument(
                new XElement(nsCore + "PageContents",
                    new XElement(nsCore + "Shapes",
                        new XElement(nsCore + "Shape",
                            new XAttribute("ID", 1),
                            new XAttribute("NameU", "Rectangle"),
                            new XAttribute("Name", "Rectangle"),
                            new XAttribute("Type", "Shape"),
                            new XElement(nsCore + "Cell", new XAttribute("N", "LineWeight"), new XAttribute("V", "0.003472222222222222")),
                            new XElement(nsCore + "Cell", new XAttribute("N", "ObjType"), new XAttribute("V", "1")),
                            new XElement(nsCore + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                                new XElement(nsCore + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "1"),
                                    new XElement(nsCore + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                                    new XElement(nsCore + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                                new XElement(nsCore + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "2"),
                                    new XElement(nsCore + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                                    new XElement(nsCore + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                                new XElement(nsCore + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "3"),
                                    new XElement(nsCore + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                                    new XElement(nsCore + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                                new XElement(nsCore + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "4"),
                                    new XElement(nsCore + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                                    new XElement(nsCore + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                                new XElement(nsCore + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "5"),
                                    new XElement(nsCore + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                                    new XElement(nsCore + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")))),
                            new XElement(nsCore + "Text", "Sample text")))));
            SaveXml(doc, path);
        }
    }
}

