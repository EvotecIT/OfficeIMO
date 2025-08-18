using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a Visio document containing pages.
    /// </summary>
    public class VisioDocument {
        private readonly List<VisioPage> _pages = new();

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
        /// Saves the document to a <c>.vsdx</c> package.
        /// </summary>
        public void Save(string filePath) {
            using Package package = Package.Open(filePath, FileMode.Create);

            Uri documentUri = new("/visio/document.xml", UriKind.Relative);
            PackagePart documentPart = package.CreatePart(documentUri, "application/vnd.ms-visio.document.main+xml");
            package.CreateRelationship(documentUri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");

            Uri pagesUri = new("/visio/pages/pages.xml", UriKind.Relative);
            PackagePart pagesPart = package.CreatePart(pagesUri, "application/vnd.ms-visio.pages+xml");
            documentPart.CreateRelationship(new Uri("pages/pages.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/pages");

            Uri page1Uri = new("/visio/pages/page1.xml", UriKind.Relative);
            PackagePart page1Part = package.CreatePart(page1Uri, "application/vnd.ms-visio.page+xml");
            PackageRelationship pageRel = pagesPart.CreateRelationship(new Uri("page1.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/page");

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
                writer.WriteStartElement("Page", ns);
                writer.WriteAttributeString("ID", "0");
                writer.WriteAttributeString("Name", pageName);
                writer.WriteAttributeString("RelId", pageRel.Id);
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

