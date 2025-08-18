using System;
using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioLoadRelId {
        [Fact]
        public void LoadIgnoresDeprecatedRelIdAttribute() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 2, 3, 4, "Rectangle"));
            document.Save(filePath);

            using (Package package = Package.Open(filePath, FileMode.Open, FileAccess.ReadWrite)) {
                PackagePart pagesPart = package.GetPart(new Uri("/visio/pages/pages.xml", UriKind.Relative));
                XDocument pagesDoc;
                using (Stream readStream = pagesPart.GetStream()) {
                    pagesDoc = XDocument.Load(readStream);
                }

                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement? pageElement = pagesDoc.Root?.Element(ns + "Page");
                pageElement?.SetAttributeValue("RelId", "rId999");

                using Stream partStream = pagesPart.GetStream(FileMode.Create, FileAccess.Write);
                pagesDoc.Save(partStream);
            }

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Single(loaded.Pages);
            VisioPage loadedPage = loaded.Pages[0];
            Assert.Single(loadedPage.Shapes);
            VisioShape shape = loadedPage.Shapes[0];
            Assert.Equal("Rectangle", shape.Text);
        }
    }
}

