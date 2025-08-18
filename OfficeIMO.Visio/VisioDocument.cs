using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
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

        private const string DocumentRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/document";
        private const string DocumentContentType = "application/vnd.ms-visio.drawing.main+xml";

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

            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XDocument pagesDoc = XDocument.Load(pagesPart.GetStream());

            foreach (XElement pageRef in pagesDoc.Root?.Elements(ns + "Page") ?? Enumerable.Empty<XElement>()) {
                string name = pageRef.Attribute("NameU")?.Value ?? pageRef.Attribute("Name")?.Value ?? "Page";
                VisioPage page = document.AddPage(name);

                string? relId = pageRef.Element(ns + "Rel")?.Attribute(rNs + "id")?.Value ?? pageRef.Attribute("RelId")?.Value;
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

        /// <summary>
        /// Saves the document to a <c>.vsdx</c> package.
        /// </summary>
        public void Save(string filePath) {
            using Package package = Package.Open(filePath, FileMode.Create);

            int relIdCounter = 1;

            Uri documentUri = new("/visio/document.xml", UriKind.Relative);
            PackagePart documentPart = package.CreatePart(documentUri, DocumentContentType);
            package.CreateRelationship(documentUri, TargetMode.Internal, DocumentRelationshipType, $"rId{relIdCounter++}");

            Uri pagesUri = new("/visio/pages/pages.xml", UriKind.Relative);
            PackagePart pagesPart = package.CreatePart(pagesUri, "application/vnd.ms-visio.pages+xml");
            documentPart.CreateRelationship(new Uri("pages/pages.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/pages", $"rId{relIdCounter++}");

            Uri page1Uri = new("/visio/pages/page1.xml", UriKind.Relative);
            PackagePart page1Part = package.CreatePart(page1Uri, "application/vnd.ms-visio.page+xml");
            PackageRelationship pageRel = pagesPart.CreateRelationship(new Uri("page1.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/page", $"rId{relIdCounter++}");

            XmlWriterSettings settings = new() {
                Encoding = new UTF8Encoding(false),
                CloseOutput = true,
                Indent = true,
            };
            const string ns = "http://schemas.microsoft.com/office/visio/2012/main";

            using (XmlWriter writer = XmlWriter.Create(documentPart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                writer.WriteStartDocument();
                writer.WriteStartElement("VisioDocument", ns);
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
            VisioShape shape = page.Shapes.Count > 0 ? page.Shapes[0] : new VisioShape("1", 1, 1, 2, 1, "Rectangle");

            using (XmlWriter writer = XmlWriter.Create(page1Part.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                writer.WriteStartDocument();
                writer.WriteStartElement("PageContents", ns);
                writer.WriteStartElement("Shapes", ns);
                writer.WriteStartElement("Shape", ns);
                writer.WriteAttributeString("ID", shape.Id);
                if (!string.IsNullOrEmpty(shape.NameU)) {
                    writer.WriteAttributeString("NameU", shape.NameU);
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
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }
    }
}

