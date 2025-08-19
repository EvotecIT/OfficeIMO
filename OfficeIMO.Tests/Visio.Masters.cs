using System;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioMasters {
        [Fact]
        public void CreatesMasterForEachNamedShape() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "A") { NameU = "Rectangle" });
            page.Shapes.Add(new VisioShape("2", 4, 1, 2, 1, "B") { NameU = "Rectangle" });
            document.Save(filePath);

            using (Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read)) {
                Assert.True(package.PartExists(new Uri("/visio/masters/masters.xml", UriKind.Relative)));
                Assert.True(package.PartExists(new Uri("/visio/masters/master1.xml", UriKind.Relative)));
                Assert.True(package.PartExists(new Uri("/visio/masters/master2.xml", UriKind.Relative)));

                PackagePart documentPart = package.GetPart(new Uri("/visio/document.xml", UriKind.Relative));
                PackageRelationship mastersRel = documentPart.GetRelationshipsByType("http://schemas.microsoft.com/visio/2010/relationships/masters").Single();
                Uri mastersUri = PackUriHelper.ResolvePartUri(documentPart.Uri, mastersRel.TargetUri);
                Assert.Equal("/visio/masters/masters.xml", mastersUri.OriginalString);

                PackagePart pagePart = package.GetPart(new Uri("/visio/pages/page1.xml", UriKind.Relative));
                PackageRelationship[] pageMasterRels = pagePart.GetRelationshipsByType("http://schemas.microsoft.com/visio/2010/relationships/master").ToArray();
                Assert.Equal(2, pageMasterRels.Length);
                Assert.Contains(pageMasterRels, r => r.TargetUri.OriginalString == "../masters/master1.xml");
                Assert.Contains(pageMasterRels, r => r.TargetUri.OriginalString == "../masters/master2.xml");

                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XDocument mastersDoc = XDocument.Load(package.GetPart(mastersUri).GetStream());
                XElement[] masterElements = mastersDoc.Root?.Elements(ns + "Master").ToArray() ?? Array.Empty<XElement>();
                Assert.Equal(2, masterElements.Length);
                Assert.All(masterElements, m => Assert.Equal("Rectangle", m.Attribute("NameU")?.Value));

                XDocument pageDoc = XDocument.Load(pagePart.GetStream());
                XElement[] shapes = pageDoc.Root?.Element(ns + "Shapes")?.Elements(ns + "Shape").ToArray() ?? Array.Empty<XElement>();
                Assert.Equal("1", shapes[0].Attribute("Master")?.Value);
                Assert.Equal("2", shapes[1].Attribute("Master")?.Value);

                PackagePart masterPart1 = package.GetPart(new Uri("/visio/masters/master1.xml", UriKind.Relative));
                XDocument master1Doc = XDocument.Load(masterPart1.GetStream());
                XElement? masterShape = master1Doc.Root?.Element(ns + "Shapes")?.Element(ns + "Shape");
                Assert.NotNull(masterShape);
                Assert.NotNull(masterShape?.Attribute("Name"));
                Assert.NotNull(masterShape?.Attribute("NameU"));
                Assert.Equal("Shape", masterShape?.Attribute("Type")?.Value);
                Assert.NotNull(masterShape?.Elements(ns + "Cell").FirstOrDefault(e => e.Attribute("N")?.Value == "PinX"));
            }

            using FileStream zipStream = File.OpenRead(filePath);
            using ZipArchive archive = new(zipStream, ZipArchiveMode.Read);
            ZipArchiveEntry entry = archive.GetEntry("[Content_Types].xml")!;
            using Stream entryStream = entry.Open();
            XDocument contentTypes = XDocument.Load(entryStream);
            XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
            Assert.NotNull(contentTypes.Root?.Elements(ct + "Override").FirstOrDefault(e => e.Attribute("PartName")?.Value == "/visio/masters/masters.xml" && e.Attribute("ContentType")?.Value == "application/vnd.ms-visio.masters+xml"));
            Assert.NotNull(contentTypes.Root?.Elements(ct + "Override").FirstOrDefault(e => e.Attribute("PartName")?.Value == "/visio/masters/master1.xml" && e.Attribute("ContentType")?.Value == "application/vnd.ms-visio.master+xml"));
            Assert.NotNull(contentTypes.Root?.Elements(ct + "Override").FirstOrDefault(e => e.Attribute("PartName")?.Value == "/visio/masters/master2.xml" && e.Attribute("ContentType")?.Value == "application/vnd.ms-visio.master+xml"));
        }
    }
}