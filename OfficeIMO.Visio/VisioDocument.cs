using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a Visio document containing pages.
    /// </summary>
    public class VisioDocument {
        private readonly List<VisioPage> _pages = new();
        private bool _requestRecalcOnOpen;

        private const string DocumentRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/document";
        private const string DocumentContentType = "application/vnd.ms-visio.drawing.main+xml";
        private const string VisioNamespace = "http://schemas.microsoft.com/office/visio/2012/main";
        private const string ThemeRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/theme";
        private const string ThemeContentType = "application/vnd.ms-visio.theme+xml";
        private const string WindowsRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/windows";
        private const string WindowsContentType = "application/vnd.ms-visio.windows+xml";

        /// <summary>
        /// Collection of pages in the document.
        /// </summary>
        public IReadOnlyList<VisioPage> Pages => _pages;

        /// <summary>
        /// Adds a new page to the document.
        /// </summary>
        /// <param name="name">Name of the page.</param>
        /// <param name="id">Optional page identifier. If not specified, uses zero-based index.</param>
        public VisioPage AddPage(string name, int? id = null) {
            VisioPage page = new(name) { Id = id ?? _pages.Count };
            _pages.Add(page);
            return page;
        }

        /// <summary>
        /// Requests Visio to relayout and reroute connectors when the document is opened.
        /// </summary>
        public void RequestRecalcOnOpen() {
            _requestRecalcOnOpen = true;
        }

        /// <summary>
        /// Loads an existing <c>.vsdx</c> file into a <see cref="VisioDocument"/>.
        /// </summary>
        /// <param name="filePath">Path to the <c>.vsdx</c> file.</param>
        public static VisioDocument Load(string filePath) {
            VisioDocument document = new();

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);

            PackageRelationship documentRel = package.GetRelationshipsByType(DocumentRelationshipType).Single();
            Uri documentUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), documentRel.TargetUri);
            PackagePart documentPart = package.GetPart(documentUri);
            if (documentPart.ContentType != DocumentContentType) {
                throw new InvalidDataException($"Unexpected Visio document content type: {documentPart.ContentType}");
            }

            PackageRelationship pagesRel = documentPart.GetRelationshipsByType("http://schemas.microsoft.com/visio/2010/relationships/pages").Single();
            Uri pagesUri = PackUriHelper.ResolvePartUri(documentPart.Uri, pagesRel.TargetUri);
            PackagePart pagesPart = package.GetPart(pagesUri);

            XNamespace ns = VisioNamespace;
            XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XDocument pagesDoc = XDocument.Load(pagesPart.GetStream());

            foreach (XElement pageRef in pagesDoc.Root?.Elements(ns + "Page") ?? Enumerable.Empty<XElement>()) {
                string name = pageRef.Attribute("Name")?.Value ?? "Page";
                int pageId = int.TryParse(pageRef.Attribute("ID")?.Value, out int tmp) ? tmp : document.Pages.Count;
                VisioPage page = document.AddPage(name, pageId);
                page.NameU = pageRef.Attribute("NameU")?.Value ?? name;
                page.ViewScale = ParseDouble(pageRef.Attribute("ViewScale")?.Value);
                page.ViewCenterX = ParseDouble(pageRef.Attribute("ViewCenterX")?.Value);
                page.ViewCenterY = ParseDouble(pageRef.Attribute("ViewCenterY")?.Value);

                string? relId = pageRef.Element(ns + "Rel")?.Attribute(rNs + "id")?.Value;
                if (string.IsNullOrEmpty(relId)) {
                    continue;
                }

                PackageRelationship pageRel = pagesPart.GetRelationship(relId);
                Uri pageUri = PackUriHelper.ResolvePartUri(pagesPart.Uri, pageRel.TargetUri);
                PackagePart pagePart = package.GetPart(pageUri);
                XDocument pageDoc = XDocument.Load(pagePart.GetStream());

                foreach (XElement shapeElement in pageDoc.Root?.Element(ns + "Shapes")?.Elements(ns + "Shape") ?? Enumerable.Empty<XElement>()) {
                    string id = shapeElement.Attribute("ID")?.Value ?? string.Empty;
                    VisioShape shape = new(id) {
                        Name = shapeElement.Attribute("Name")?.Value,
                        NameU = shapeElement.Attribute("NameU")?.Value,
                        Text = shapeElement.Element(ns + "Text")?.Value
                    };

                    var cellElements = shapeElement.Elements(ns + "Cell").ToList();
                    if (cellElements.Count > 0) {
                        foreach (XElement cell in cellElements) {
                            string? n = cell.Attribute("N")?.Value;
                            string? v = cell.Attribute("V")?.Value;
                            switch (n) {
                                case "PinX":
                                    shape.PinX = ParseDouble(v);
                                    break;
                                case "PinY":
                                    shape.PinY = ParseDouble(v);
                                    break;
                                case "Width":
                                    shape.Width = ParseDouble(v);
                                    break;
                                case "Height":
                                    shape.Height = ParseDouble(v);
                                    break;
                            }
                        }
                    } else {
                        XElement? xform = shapeElement.Element(ns + "XForm");
                        shape.PinX = ParseDouble(xform?.Element(ns + "PinX")?.Value);
                        shape.PinY = ParseDouble(xform?.Element(ns + "PinY")?.Value);
                        shape.Width = ParseDouble(xform?.Element(ns + "Width")?.Value);
                        shape.Height = ParseDouble(xform?.Element(ns + "Height")?.Value);
                    }

                    page.Shapes.Add(shape);
                }
            }

            return document;
        }

        private static double ParseDouble(string? value) {
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double result) ? result : 0;
        }

        private static XDocument CreateVisioDocumentXml(bool requestRecalcOnOpen) {
            XNamespace ns = VisioNamespace;
            XElement settings = new(ns + "DocumentSettings");
            if (requestRecalcOnOpen) {
                settings.Add(new XElement(ns + "RelayoutAndRerouteUponOpen", 1));
            }

            return new XDocument(
                new XElement(ns + "VisioDocument",
                    settings,
                    new XElement(ns + "Colors"),
                    new XElement(ns + "FaceNames"),
                    new XElement(ns + "StyleSheets")));
        }

        /// <summary>
        /// Saves the document to a <c>.vsdx</c> package.
        /// </summary>
        public void Save(string filePath) {
            int masterCount = 0;
            bool includeTheme = _pages.Any(p => p.Shapes.Any());
            using (Package package = Package.Open(filePath, FileMode.Create)) {
                Uri documentUri = new("/visio/document.xml", UriKind.Relative);
                PackagePart documentPart = package.CreatePart(documentUri, DocumentContentType);
                package.CreateRelationship(documentUri, TargetMode.Internal, DocumentRelationshipType, "rId1");

                Uri coreUri = new("/docProps/core.xml", UriKind.Relative);
                PackagePart corePart = package.CreatePart(coreUri, "application/vnd.openxmlformats-package.core-properties+xml");
                package.CreateRelationship(coreUri, TargetMode.Internal, "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "rId2");

                Uri appUri = new("/docProps/app.xml", UriKind.Relative);
                PackagePart appPart = package.CreatePart(appUri, "application/vnd.openxmlformats-officedocument.extended-properties+xml");
                package.CreateRelationship(appUri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "rId3");

                Uri customUri = new("/docProps/custom.xml", UriKind.Relative);
                PackagePart customPart = package.CreatePart(customUri, "application/vnd.openxmlformats-officedocument.custom-properties+xml");
                package.CreateRelationship(customUri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties", "rId4");

                Uri thumbUri = new("/docProps/thumbnail.emf", UriKind.Relative);
                PackagePart thumbPart = package.CreatePart(thumbUri, "image/x-emf");
                package.CreateRelationship(thumbUri, TargetMode.Internal, "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail", "rId5");

                Uri pagesUri = new("/visio/pages/pages.xml", UriKind.Relative);
                PackagePart pagesPart = package.CreatePart(pagesUri, "application/vnd.ms-visio.pages+xml");
                documentPart.CreateRelationship(new Uri("pages/pages.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/pages", "rId1");

                Uri windowsUri = new("/visio/windows.xml", UriKind.Relative);
                PackagePart windowsPart = package.CreatePart(windowsUri, WindowsContentType);
                documentPart.CreateRelationship(new Uri("windows.xml", UriKind.Relative), TargetMode.Internal, WindowsRelationshipType, "rId2");

                PackagePart? themePart = null;
                if (includeTheme) {
                    Uri themeUri = new("/visio/theme/theme1.xml", UriKind.Relative);
                    themePart = package.CreatePart(themeUri, ThemeContentType);
                    documentPart.CreateRelationship(new Uri("theme/theme1.xml", UriKind.Relative), TargetMode.Internal, ThemeRelationshipType, "rId3");
                }

                Uri page1Uri = new("/visio/pages/page1.xml", UriKind.Relative);
                PackagePart page1Part = package.CreatePart(page1Uri, "application/vnd.ms-visio.page+xml");
                PackageRelationship pageRel = pagesPart.CreateRelationship(new Uri("page1.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/page", "rId1");

                XmlWriterSettings settings = new() {
                    Encoding = new UTF8Encoding(false),
                    CloseOutput = true,
                    Indent = false,
                };
                const string ns = VisioNamespace;

                void WriteCell(XmlWriter writer, string name, double value) {
                    writer.WriteStartElement("Cell", ns);
                    writer.WriteAttributeString("N", name);
                    writer.WriteAttributeString("V", value.ToString(CultureInfo.InvariantCulture));
                    writer.WriteEndElement();
                }

                if (themePart != null) {
                    using (XmlWriter writer = XmlWriter.Create(themePart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                        writer.WriteStartDocument();
                        writer.WriteStartElement("a", "theme", "http://schemas.openxmlformats.org/drawingml/2006/main");
                        writer.WriteAttributeString("name", "Office Theme");
                        writer.WriteEndElement();
                        writer.WriteEndDocument();
                    }
                }

                List<VisioMaster> masters = new();
                foreach (VisioShape shape in _pages.SelectMany(p => p.Shapes).Where(s => !string.IsNullOrEmpty(s.NameU))) {
                    VisioMaster master = new((masters.Count + 2).ToString(CultureInfo.InvariantCulture), shape.NameU!, shape);
                    masters.Add(master);
                    shape.Master = master;
                }

                PackagePart? mastersPart = null;
                if (masters.Count > 0) {
                    Uri mastersUri = new("/visio/masters/masters.xml", UriKind.Relative);
                    mastersPart = package.CreatePart(mastersUri, "application/vnd.ms-visio.masters+xml");
                    documentPart.CreateRelationship(new Uri("masters/masters.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/masters", "rId4");

                    for (int i = 0; i < masters.Count; i++) {
                        VisioMaster master = masters[i];
                        Uri masterUri = new($"/visio/masters/master{i + 1}.xml", UriKind.Relative);
                        PackagePart masterPart = package.CreatePart(masterUri, "application/vnd.ms-visio.master+xml");
                        mastersPart.CreateRelationship(new Uri($"master{i + 1}.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/master", $"rId{i + 1}");
                        page1Part.CreateRelationship(new Uri($"../masters/master{i + 1}.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/master", $"rId{i + 1}");

                        using (XmlWriter writer = XmlWriter.Create(masterPart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                            writer.WriteStartDocument();
                            writer.WriteStartElement("MasterContents", ns);
                            writer.WriteStartElement("Shapes", ns);
                            VisioShape s = master.Shape;
                            writer.WriteStartElement("Shape", ns);
                            writer.WriteAttributeString("ID", "1");
                            string masterShapeName = s.Name ?? s.NameU ?? "MasterShape";
                            writer.WriteAttributeString("Name", masterShapeName);
                            writer.WriteAttributeString("NameU", master.NameU);
                            writer.WriteAttributeString("Type", "Shape");
                            WriteCell(writer, "PinX", s.PinX);
                            WriteCell(writer, "PinY", s.PinY);
                            WriteCell(writer, "Width", s.Width);
                            WriteCell(writer, "Height", s.Height);
                            if (!string.IsNullOrEmpty(s.Text)) {
                                writer.WriteElementString("Text", ns, s.Text);
                            }
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                        }
                    }

                    using (XmlWriter writer = XmlWriter.Create(mastersPart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                        writer.WriteStartDocument();
                        writer.WriteStartElement("Masters", ns);
                        writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                        for (int i = 0; i < masters.Count; i++) {
                            VisioMaster master = masters[i];
                            writer.WriteStartElement("Master", ns);
                            writer.WriteAttributeString("ID", master.Id);
                            writer.WriteAttributeString("NameU", master.NameU);
                            writer.WriteStartElement("Rel", ns);
                            writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", $"rId{i + 1}");
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement();
                        writer.WriteEndDocument();
                    }
                }

                using (Stream stream = documentPart.GetStream(FileMode.Create, FileAccess.Write)) {
                    CreateVisioDocumentXml(_requestRecalcOnOpen).Save(stream);
                }

                using (XmlWriter writer = XmlWriter.Create(corePart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("cp", "coreProperties", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
                    writer.WriteAttributeString("xmlns", "dc", null, "http://purl.org/dc/elements/1.1/");
                    writer.WriteAttributeString("xmlns", "dcterms", null, "http://purl.org/dc/terms/");
                    writer.WriteAttributeString("xmlns", "dcmitype", null, "http://purl.org/dc/dcmitype/");
                    writer.WriteAttributeString("xmlns", "xsi", null, "http://www.w3.org/2001/XMLSchema-instance");
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }

                using (XmlWriter writer = XmlWriter.Create(appPart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("ep", "Properties", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
                    writer.WriteAttributeString("xmlns", "vt", null, "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }

                using (XmlWriter writer = XmlWriter.Create(customPart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("cp", "Properties", "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties");
                    writer.WriteAttributeString("xmlns", "vt", null, "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }

                using (Stream stream = thumbPart.GetStream(FileMode.Create, FileAccess.Write)) { }

                using (XmlWriter writer = XmlWriter.Create(windowsPart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("Windows", ns);
                    writer.WriteAttributeString("ClientWidth", "1000");
                    writer.WriteAttributeString("ClientHeight", "1000");
                    writer.WriteStartElement("Window", ns);
                    writer.WriteAttributeString("WindowType", "1");
                    writer.WriteAttributeString("WindowState", "0");
                    writer.WriteAttributeString("ClientLeft", "0");
                    writer.WriteAttributeString("ClientTop", "0");
                    writer.WriteAttributeString("ClientWidth", "1000");
                    writer.WriteAttributeString("ClientHeight", "1000");
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }

                VisioPage page = _pages.Count > 0 ? _pages[0] : new VisioPage("Page-1");

                using (XmlWriter writer = XmlWriter.Create(pagesPart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("Pages", ns);
                    writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    writer.WriteAttributeString("xml", "space", null, "preserve");
                    writer.WriteStartElement("Page", ns);
                    writer.WriteAttributeString("ID", page.Id.ToString(CultureInfo.InvariantCulture));
                    writer.WriteAttributeString("Name", page.Name);
                    writer.WriteAttributeString("NameU", page.NameU ?? page.Name);
                    writer.WriteAttributeString("ViewScale", page.ViewScale.ToString(CultureInfo.InvariantCulture));
                    writer.WriteAttributeString("ViewCenterX", page.ViewCenterX.ToString(CultureInfo.InvariantCulture));
                    writer.WriteAttributeString("ViewCenterY", page.ViewCenterY.ToString(CultureInfo.InvariantCulture));
                    writer.WriteStartElement("PageSheet", ns);
                    writer.WriteAttributeString("LineStyle", "0");
                    writer.WriteAttributeString("FillStyle", "0");
                    writer.WriteAttributeString("TextStyle", "0");
                    void WritePageCell(string name, double value, string? unit = null, string? formula = null) {
                        writer.WriteStartElement("Cell", ns);
                        writer.WriteAttributeString("N", name);
                        writer.WriteAttributeString("V", value.ToString(CultureInfo.InvariantCulture));
                        if (unit != null) writer.WriteAttributeString("U", unit);
                        if (formula != null) writer.WriteAttributeString("F", formula);
                        writer.WriteEndElement();
                    }
                    bool useUnits = page.PageWidth != 8.26771653543307 || page.PageHeight != 11.69291338582677;
                    WritePageCell("PageWidth", page.PageWidth, useUnits ? "MM" : null);
                    WritePageCell("PageHeight", page.PageHeight, useUnits ? "MM" : null);
                    WritePageCell("ShdwOffsetX", 0.1181102362204724, useUnits ? "MM" : null);
                    WritePageCell("ShdwOffsetY", -0.1181102362204724, useUnits ? "MM" : null);
                    WritePageCell("PageScale", 0.03937007874015748, "MM");
                    WritePageCell("DrawingScale", 0.03937007874015748, "MM");
                    WritePageCell("DrawingSizeType", 0);
                    WritePageCell("DrawingScaleType", 0);
                    WritePageCell("InhibitSnap", 0);
                    WritePageCell("PageLockReplace", 0, "BOOL");
                    WritePageCell("PageLockDuplicate", 0, "BOOL");
                    WritePageCell("UIVisibility", 0);
                    WritePageCell("ShdwType", 0);
                    WritePageCell("ShdwObliqueAngle", 0);
                    WritePageCell("ShdwScaleFactor", 1);
                    WritePageCell("DrawingResizeType", 1);
                    WritePageCell("PageShapeSplit", 1);
                    if (includeTheme) {
                        WritePageCell("ColorSchemeIndex", 60);
                        WritePageCell("EffectSchemeIndex", 60);
                        WritePageCell("ConnectorSchemeIndex", 60);
                        WritePageCell("FontSchemeIndex", 60);
                        WritePageCell("ThemeIndex", 60);
                        WritePageCell("PageLeftMargin", 0.25, "MM");
                        WritePageCell("PageRightMargin", 0.25, "MM");
                        WritePageCell("PageTopMargin", 0.25, "MM");
                        WritePageCell("PageBottomMargin", 0.25, "MM");
                        WritePageCell("PrintPageOrientation", 2);
                        writer.WriteStartElement("Section", ns);
                        writer.WriteAttributeString("N", "User");
                        writer.WriteStartElement("Row", ns);
                        writer.WriteAttributeString("N", "msvThemeOrder");
                        writer.WriteStartElement("Cell", ns);
                        writer.WriteAttributeString("N", "Value");
                        writer.WriteAttributeString("V", "0");
                        writer.WriteEndElement();
                        writer.WriteStartElement("Cell", ns);
                        writer.WriteAttributeString("N", "Prompt");
                        writer.WriteAttributeString("V", "");
                        writer.WriteAttributeString("F", "No Formula");
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                    writer.WriteStartElement("Rel", ns);
                    writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", pageRel.Id);
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }

                using (XmlWriter writer = XmlWriter.Create(page1Part.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("PageContents", ns);
                    writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    writer.WriteAttributeString("xml", "space", null, "preserve");
                    if (page.Shapes.Count > 0 || page.Connectors.Count > 0) {
                        writer.WriteStartElement("Shapes", ns);

                        foreach (VisioShape shape in page.Shapes) {
                            writer.WriteStartElement("Shape", ns);
                            writer.WriteAttributeString("ID", shape.Id);
                            string shapeName = shape.Name ?? shape.NameU ?? $"Shape{shape.Id}";
                            writer.WriteAttributeString("Name", shapeName);
                            writer.WriteAttributeString("NameU", shape.NameU ?? shapeName);
                            writer.WriteAttributeString("Type", "Shape");
                            if (shape.Master != null) {
                                writer.WriteAttributeString("Master", shape.Master.Id);
                                WriteCell(writer, "PinX", shape.PinX);
                                WriteCell(writer, "PinY", shape.PinY);
                                writer.WriteStartElement("Cell", ns);
                                writer.WriteAttributeString("N", "LineWeight");
                                writer.WriteAttributeString("V", "0.003472222222222222");
                                writer.WriteAttributeString("U", "PT");
                                writer.WriteAttributeString("F", "Inh");
                                writer.WriteEndElement();
                                if (!string.IsNullOrEmpty(shape.Text)) {
                                    writer.WriteElementString("Text", ns, shape.Text);
                                }
                            } else {
                                WriteCell(writer, "PinX", shape.PinX);
                                WriteCell(writer, "PinY", shape.PinY);
                                WriteCell(writer, "Width", shape.Width);
                                WriteCell(writer, "Height", shape.Height);
                                if (!string.IsNullOrEmpty(shape.Text)) {
                                    writer.WriteElementString("Text", ns, shape.Text);
                                }
                            }
                            writer.WriteEndElement();
                        }

                        foreach (VisioConnector connector in page.Connectors) {
                            VisioShape from = connector.From;
                            VisioShape to = connector.To;
                            double startX = from.PinX + from.Width / 2;
                            double startY = from.PinY;
                            double endX = to.PinX - to.Width / 2;
                            double endY = to.PinY;

                            writer.WriteStartElement("Shape", ns);
                            writer.WriteAttributeString("ID", connector.Id);
                            writer.WriteAttributeString("Name", "Connector");
                            writer.WriteAttributeString("NameU", "Connector");
                            writer.WriteAttributeString("Type", "Shape");
                            writer.WriteStartElement("Geom", ns);
                            writer.WriteStartElement("MoveTo", ns);
                            writer.WriteAttributeString("X", startX.ToString(CultureInfo.InvariantCulture));
                            writer.WriteAttributeString("Y", startY.ToString(CultureInfo.InvariantCulture));
                            writer.WriteEndElement();
                            writer.WriteStartElement("LineTo", ns);
                            writer.WriteAttributeString("X", startX.ToString(CultureInfo.InvariantCulture));
                            writer.WriteAttributeString("Y", endY.ToString(CultureInfo.InvariantCulture));
                            writer.WriteEndElement();
                            writer.WriteStartElement("LineTo", ns);
                            writer.WriteAttributeString("X", endX.ToString(CultureInfo.InvariantCulture));
                            writer.WriteAttributeString("Y", endY.ToString(CultureInfo.InvariantCulture));
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                        }

                        writer.WriteEndElement(); // Shapes

                        if (page.Connectors.Count > 0) {
                            writer.WriteStartElement("Connects", ns);
                            foreach (VisioConnector connector in page.Connectors) {
                                writer.WriteStartElement("Connect", ns);
                                writer.WriteAttributeString("FromSheet", connector.Id);
                                writer.WriteAttributeString("FromCell", "BeginX");
                                writer.WriteAttributeString("ToSheet", connector.From.Id);
                                writer.WriteAttributeString("ToCell", "PinX");
                                writer.WriteEndElement();
                                writer.WriteStartElement("Connect", ns);
                                writer.WriteAttributeString("FromSheet", connector.Id);
                                writer.WriteAttributeString("FromCell", "EndX");
                                writer.WriteAttributeString("ToSheet", connector.To.Id);
                                writer.WriteAttributeString("ToCell", "PinX");
                                writer.WriteEndElement();
                            }
                            writer.WriteEndElement(); // Connects
                        }
                    }

                    writer.WriteEndElement(); // PageContents
                    writer.WriteEndDocument();
                }
                masterCount = masters.Count;
            }

            FixContentTypes(filePath, masterCount, includeTheme);
        }

        private static void FixContentTypes(string filePath, int masterCount, bool includeTheme) {
            using FileStream zipStream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite);
            using ZipArchive archive = new(zipStream, ZipArchiveMode.Update);
            ZipArchiveEntry? entry = archive.GetEntry("[Content_Types].xml");
            entry?.Delete();
            ZipArchiveEntry newEntry = archive.CreateEntry("[Content_Types].xml");
            XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
            XElement root = new(ct + "Types",
                new XElement(ct + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                new XElement(ct + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")),
                new XElement(ct + "Default", new XAttribute("Extension", "emf"), new XAttribute("ContentType", "image/x-emf")),
                new XElement(ct + "Override", new XAttribute("PartName", "/visio/document.xml"), new XAttribute("ContentType", "application/vnd.ms-visio.drawing.main+xml")),
                new XElement(ct + "Override", new XAttribute("PartName", "/visio/pages/pages.xml"), new XAttribute("ContentType", "application/vnd.ms-visio.pages+xml")),
                new XElement(ct + "Override", new XAttribute("PartName", "/visio/pages/page1.xml"), new XAttribute("ContentType", "application/vnd.ms-visio.page+xml")),
                new XElement(ct + "Override", new XAttribute("PartName", "/docProps/core.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.core-properties+xml")),
                new XElement(ct + "Override", new XAttribute("PartName", "/docProps/app.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml")),
                new XElement(ct + "Override", new XAttribute("PartName", "/docProps/custom.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.custom-properties+xml")),
                new XElement(ct + "Override", new XAttribute("PartName", "/docProps/thumbnail.emf"), new XAttribute("ContentType", "image/x-emf")),
                new XElement(ct + "Override", new XAttribute("PartName", "/visio/windows.xml"), new XAttribute("ContentType", WindowsContentType)));
            if (includeTheme) {
                root.Add(new XElement(ct + "Override", new XAttribute("PartName", "/visio/theme/theme1.xml"), new XAttribute("ContentType", ThemeContentType)));
            }
            if (masterCount > 0) {
                root.Add(new XElement(ct + "Override", new XAttribute("PartName", "/visio/masters/masters.xml"), new XAttribute("ContentType", "application/vnd.ms-visio.masters+xml")));
                for (int i = 1; i <= masterCount; i++) {
                    root.Add(new XElement(ct + "Override", new XAttribute("PartName", $"/visio/masters/master{i}.xml"), new XAttribute("ContentType", "application/vnd.ms-visio.master+xml")));
                }
            }
            XDocument doc = new(root);
            using StreamWriter writer = new(newEntry.Open());
            writer.Write(doc.Declaration + Environment.NewLine + doc.ToString(SaveOptions.DisableFormatting));
        }
    }
}