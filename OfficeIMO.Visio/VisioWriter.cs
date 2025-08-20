using System;
using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public static class VisioWriter {
        private static readonly XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
        private static readonly XNamespace rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private const string RT_Document = "http://schemas.microsoft.com/visio/2010/relationships/document";
        private const string RT_Pages = "http://schemas.microsoft.com/visio/2010/relationships/pages";
        private const string RT_Page = "http://schemas.microsoft.com/visio/2010/relationships/page";

        private const string CT_Document = "application/vnd.ms-visio.drawing.main+xml";
        private const string CT_Pages = "application/vnd.ms-visio.pages+xml";
        private const string CT_Page = "application/vnd.ms-visio.page+xml";

        public static void Create(string filePath) {
            if (File.Exists(filePath)) {
                File.Delete(filePath);
            }

            using Package package = Package.Open(filePath, FileMode.Create, FileAccess.ReadWrite);

            Uri documentUri = PackUriHelper.CreatePartUri(new Uri("/visio/document.xml", UriKind.Relative));
            Uri pagesUri = PackUriHelper.CreatePartUri(new Uri("/visio/pages/pages.xml", UriKind.Relative));
            Uri page1Uri = PackUriHelper.CreatePartUri(new Uri("/visio/pages/page1.xml", UriKind.Relative));

            PackagePart documentPart = package.CreatePart(documentUri, CT_Document, CompressionOption.Maximum);
            PackagePart pagesPart = package.CreatePart(pagesUri, CT_Pages, CompressionOption.Maximum);
            PackagePart page1Part = package.CreatePart(page1Uri, CT_Page, CompressionOption.Maximum);

            package.CreateRelationship(documentUri, TargetMode.Internal, RT_Document, "rId1");

            documentPart.CreateRelationship(pagesUri, TargetMode.Internal, RT_Pages, "rId1");
            pagesPart.CreateRelationship(page1Uri, TargetMode.Internal, RT_Page, "rId1");

            WriteDocumentXml(documentPart.GetStream(FileMode.Create, FileAccess.Write));
            WritePagesXml(pagesPart.GetStream(FileMode.Create, FileAccess.Write));
            WritePage1Xml(page1Part.GetStream(FileMode.Create, FileAccess.Write));
        }

        private static void WriteDocumentXml(Stream stream) {
            XDocument doc = new(new XDeclaration("1.0", "utf-8", null),
                new XElement(v + "VisioDocument",
                    new XElement(v + "DocumentSettings",
                        new XElement(v + "RelayoutAndRerouteUponOpen", 1)
                    ),
                    new XElement(v + "Colors"),
                    new XElement(v + "FaceNames"),
                    new XElement(v + "StyleSheets")));
            using StreamWriter writer = new(stream);
            writer.Write(doc.Declaration + Environment.NewLine + doc.ToString(SaveOptions.DisableFormatting));
        }

        private static void WritePagesXml(Stream stream) {
            XDocument doc = new(new XDeclaration("1.0", "utf-8", null),
                new XElement(v + "Pages",
                    new XAttribute(XNamespace.Xmlns + "r", rel),
                    new XElement(v + "Page",
                        new XAttribute("ID", 1),
                        new XAttribute("Name", "Page-1"),
                        new XElement(v + "Rel",
                            new XAttribute(rel + "id", "rId1")))));
            using StreamWriter writer = new(stream);
            writer.Write(doc.Declaration + Environment.NewLine + doc.ToString(SaveOptions.DisableFormatting));
        }

        private static void WritePage1Xml(Stream stream) {
            XDocument doc = new(new XDeclaration("1.0", "utf-8", null),
                new XElement(v + "PageContents",
                    new XElement(v + "Shapes",
                        new XElement(v + "Shape",
                            new XAttribute("ID", 1),
                            new XAttribute("NameU", "Start"),
                            new XElement(v + "XForm",
                                new XElement(v + "PinX", 1.0),
                                new XElement(v + "PinY", 1.0),
                                new XElement(v + "Width", 2.0),
                                new XElement(v + "Height", 1.0),
                                new XElement(v + "LocPinX", 1.0),
                                new XElement(v + "LocPinY", 0.5),
                                new XElement(v + "Angle", 0.0)),
                            new XElement(v + "Geom",
                                new XElement(v + "MoveTo",
                                    new XAttribute("X", 0.0),
                                    new XAttribute("Y", 0.0)),
                                new XElement(v + "LineTo",
                                    new XAttribute("X", 2.0),
                                    new XAttribute("Y", 0.0)),
                                new XElement(v + "LineTo",
                                    new XAttribute("X", 2.0),
                                    new XAttribute("Y", 1.0)),
                                new XElement(v + "LineTo",
                                    new XAttribute("X", 0.0),
                                    new XAttribute("Y", 1.0)),
                                new XElement(v + "LineTo",
                                    new XAttribute("X", 0.0),
                                    new XAttribute("Y", 0.0))),
                            new XElement(v + "Cell",
                                new XAttribute("N", "LineWeight"),
                                new XAttribute("V", 0.0138889)),
                            new XElement(v + "Text", "Start")),
                        new XElement(v + "Shape",
                            new XAttribute("ID", 2),
                            new XAttribute("NameU", "End"),
                            new XElement(v + "XForm",
                                new XElement(v + "PinX", 4.0),
                                new XElement(v + "PinY", 1.0),
                                new XElement(v + "Width", 2.0),
                                new XElement(v + "Height", 1.0),
                                new XElement(v + "LocPinX", 1.0),
                                new XElement(v + "LocPinY", 0.5),
                                new XElement(v + "Angle", 0.0)),
                            new XElement(v + "Geom",
                                new XElement(v + "MoveTo",
                                    new XAttribute("X", 0.0),
                                    new XAttribute("Y", 0.0)),
                                new XElement(v + "LineTo",
                                    new XAttribute("X", 2.0),
                                    new XAttribute("Y", 0.0)),
                                new XElement(v + "LineTo",
                                    new XAttribute("X", 2.0),
                                    new XAttribute("Y", 1.0)),
                                new XElement(v + "LineTo",
                                    new XAttribute("X", 0.0),
                                    new XAttribute("Y", 1.0)),
                                new XElement(v + "LineTo",
                                    new XAttribute("X", 0.0),
                                    new XAttribute("Y", 0.0))),
                            new XElement(v + "Cell",
                                new XAttribute("N", "LineWeight"),
                                new XAttribute("V", 0.0138889)),
                            new XElement(v + "Text", "End")),
                        new XElement(v + "Shape",
                            new XAttribute("ID", 3),
                            new XAttribute("NameU", "Connector"),
                            new XElement(v + "Geom",
                                new XElement(v + "MoveTo",
                                    new XAttribute("X", 2.0),
                                    new XAttribute("Y", 1.0)),
                                new XElement(v + "LineTo",
                                    new XAttribute("X", 3.0),
                                    new XAttribute("Y", 1.0))))),
                    new XElement(v + "Connects",
                        new XElement(v + "Connect",
                            new XAttribute("FromSheet", 3),
                            new XAttribute("FromCell", "BeginX"),
                            new XAttribute("ToSheet", 1),
                            new XAttribute("ToCell", "PinX")),
                        new XElement(v + "Connect",
                            new XAttribute("FromSheet", 3),
                            new XAttribute("FromCell", "EndX"),
                            new XAttribute("ToSheet", 2),
                            new XAttribute("ToCell", "PinX")))));
            using StreamWriter writer = new(stream);
            writer.Write(doc.Declaration + Environment.NewLine + doc.ToString(SaveOptions.DisableFormatting));
        }
    }
}
