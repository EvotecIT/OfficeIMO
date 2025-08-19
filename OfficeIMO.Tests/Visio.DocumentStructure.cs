using System;
using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioDocumentStructure {
        [Fact]
        public void VisioDocumentIncludesRequiredChildren() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = new();
            document.Save(filePath);

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            PackagePart documentPart = package.GetPart(new Uri("/visio/document.xml", UriKind.Relative));
            XDocument xml = XDocument.Load(documentPart.GetStream());
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";

            Assert.NotNull(xml.Root?.Element(ns + "DocumentSettings"));
            Assert.NotNull(xml.Root?.Element(ns + "Colors"));
            Assert.NotNull(xml.Root?.Element(ns + "FaceNames"));
            Assert.NotNull(xml.Root?.Element(ns + "StyleSheets"));
        }
    }
}

