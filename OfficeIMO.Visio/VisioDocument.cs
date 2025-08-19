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

        /// <summary>
        /// Collection of pages in the document.
        /// </summary>
        public IReadOnlyList<VisioPage> Pages => _pages;

        /// <summary>
        /// Adds a new page to the document.
        /// </summary>
        /// <param name="name">Name of the page.</param>
        public VisioPage AddPage(string name) {
            VisioPage page = new(name);
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
                string name = pageRef.Attribute("NameU")?.Value ?? pageRef.Attribute("Name")?.Value ?? "Page";
                VisioPage page = document.AddPage(name);

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
                        NameU = shapeElement.Attribute("NameU")?.Value,
                        Text = shapeElement.Element(ns + "Text")?.Value
                    };

                    XElement? xform = shapeElement.Element(ns + "XForm");
                    shape.PinX = ParseDouble(xform?.Element(ns + "PinX")?.Value);
                    shape.PinY = ParseDouble(xform?.Element(ns + "PinY")?.Value);
                    shape.Width = ParseDouble(xform?.Element(ns + "Width")?.Value);
                    shape.Height = ParseDouble(xform?.Element(ns + "Height")?.Value);

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

                Uri windowsUri = new("/visio/windows.xml", UriKind.Relative);
                PackagePart windowsPart = package.CreatePart(windowsUri, "application/vnd.ms-visio.windows+xml");
                package.CreateRelationship(windowsUri, TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/windows", "rId6");

                Uri pagesUri = new("/visio/pages/pages.xml", UriKind.Relative);
                PackagePart pagesPart = package.CreatePart(pagesUri, "application/vnd.ms-visio.pages+xml");
                documentPart.CreateRelationship(new Uri("pages/pages.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/pages", "rId1");

                Uri page1Uri = new("/visio/pages/page1.xml", UriKind.Relative);
                PackagePart page1Part = package.CreatePart(page1Uri, "application/vnd.ms-visio.page+xml");
                PackageRelationship pageRel = pagesPart.CreateRelationship(new Uri("page1.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/page", "rId1");

                XmlWriterSettings settings = new() {
                    Encoding = new UTF8Encoding(false),
                    CloseOutput = true,
                    Indent = true,
                };
                const string ns = VisioNamespace;

                List<VisioMaster> masters = new();
                foreach (IGrouping<string, VisioShape> group in _pages.SelectMany(p => p.Shapes).Where(s => !string.IsNullOrEmpty(s.NameU)).GroupBy(s => s.NameU!)) {
                    if (group.Count() < 2) {
                        continue;
                    }

                    VisioShape template = group.First();
                    VisioMaster master = new((masters.Count + 1).ToString(CultureInfo.InvariantCulture), group.Key, template);
                    masters.Add(master);
                    foreach (VisioShape item in group) {
                        item.Master = master;
                    }
                }

                PackagePart? mastersPart = null;
                if (masters.Count > 0) {
                    Uri mastersUri = new("/visio/masters/masters.xml", UriKind.Relative);
                    mastersPart = package.CreatePart(mastersUri, "application/vnd.ms-visio.masters+xml");
                    documentPart.CreateRelationship(new Uri("masters/masters.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/masters", "rId2");

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
                            writer.WriteAttributeString("NameU", master.NameU);
                            writer.WriteStartElement("XForm", ns);
                            writer.WriteElementString("PinX", ns, s.PinX.ToString(CultureInfo.InvariantCulture));
                            writer.WriteElementString("PinY", ns, s.PinY.ToString(CultureInfo.InvariantCulture));
                            writer.WriteElementString("Width", ns, s.Width.ToString(CultureInfo.InvariantCulture));
                            writer.WriteElementString("Height", ns, s.Height.ToString(CultureInfo.InvariantCulture));
                            writer.WriteEndElement();
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
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }

                string pageName = _pages.Count > 0 ? _pages[0].Name : "Page-1";
                using (XmlWriter writer = XmlWriter.Create(pagesPart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("Pages", ns);
                    writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    writer.WriteStartElement("Page", ns);
                    writer.WriteAttributeString("ID", "1");
                    writer.WriteAttributeString("Name", pageName);
                    writer.WriteStartElement("Rel", ns);
                    writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", pageRel.Id);
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }

                VisioPage page = _pages.Count > 0 ? _pages[0] : new VisioPage(pageName);

                using (XmlWriter writer = XmlWriter.Create(page1Part.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("PageContents", ns);
                    writer.WriteStartElement("Shapes", ns);

                    foreach (VisioShape shape in page.Shapes) {
                        writer.WriteStartElement("Shape", ns);
                        writer.WriteAttributeString("ID", shape.Id);
                        if (!string.IsNullOrEmpty(shape.NameU)) {
                            writer.WriteAttributeString("NameU", shape.NameU);
                        }
                        if (shape.Master != null) {
                            writer.WriteAttributeString("Master", shape.Master.Id);
                        }
                        writer.WriteStartElement("XForm", ns);
                        writer.WriteElementString("PinX", ns, shape.PinX.ToString(CultureInfo.InvariantCulture));
                        writer.WriteElementString("PinY", ns, shape.PinY.ToString(CultureInfo.InvariantCulture));
                        writer.WriteElementString("Width", ns, shape.Width.ToString(CultureInfo.InvariantCulture));
                        writer.WriteElementString("Height", ns, shape.Height.ToString(CultureInfo.InvariantCulture));
                        writer.WriteEndElement();
                        if (!string.IsNullOrEmpty(shape.Text)) {
                            writer.WriteElementString("Text", ns, shape.Text);
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
                        writer.WriteAttributeString("NameU", "Connector");
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

                    writer.WriteEndElement(); // PageContents
                    writer.WriteEndDocument();
                }
                masterCount = masters.Count;
            }

            FixContentTypes(filePath, masterCount);
        }

        private static void FixContentTypes(string filePath, int masterCount) {
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
                new XElement(ct + "Override", new XAttribute("PartName", "/visio/windows.xml"), new XAttribute("ContentType", "application/vnd.ms-visio.windows+xml")));
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