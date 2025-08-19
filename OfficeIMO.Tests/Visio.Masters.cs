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
        public void ReusesMasterForDuplicateShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "A") { NameU = "Rectangle" });
            page.Shapes.Add(new VisioShape("2", 4, 1, 2, 1, "B") { NameU = "Rectangle" });
            document.Save(filePath);

            using (Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read)) {
                Assert.True(package.PartExists(new Uri("/visio/masters/masters.xml", UriKind.Relative)));
                Assert.True(package.PartExists(new Uri("/visio/masters/master1.xml", UriKind.Relative)));

                PackagePart documentPart = package.GetPart(new Uri("/visio/document.xml", UriKind.Relative));
                PackageRelationship mastersRel = documentPart.GetRelationshipsByType("http://schemas.microsoft.com/visio/2010/relationships/masters").Single();
                Uri mastersUri = PackUriHelper.ResolvePartUri(documentPart.Uri, mastersRel.TargetUri);
                Assert.Equal("/visio/masters/masters.xml", mastersUri.OriginalString);

                PackagePart pagePart = package.GetPart(new Uri("/visio/pages/page1.xml", UriKind.Relative));
                PackageRelationship pageMasterRel = pagePart.GetRelationshipsByType("http://schemas.microsoft.com/visio/2010/relationships/master").Single();
                Assert.Equal("../masters/master1.xml", pageMasterRel.TargetUri.OriginalString);

                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XDocument mastersDoc = XDocument.Load(package.GetPart(mastersUri).GetStream());
                XElement master = mastersDoc.Root?.Element(ns + "Master");
                Assert.Equal("1", master?.Attribute("ID")?.Value);
                Assert.Equal("Rectangle", master?.Attribute("NameU")?.Value);

                XDocument pageDoc = XDocument.Load(pagePart.GetStream());
                XElement shape = pageDoc.Root?.Element(ns + "Shapes")?.Element(ns + "Shape");
                Assert.Equal("1", shape?.Attribute("Master")?.Value);
            }

            using FileStream zipStream = File.OpenRead(filePath);
            using ZipArchive archive = new(zipStream, ZipArchiveMode.Read);
            ZipArchiveEntry entry = archive.GetEntry("[Content_Types].xml")!;
            using Stream entryStream = entry.Open();
            XDocument contentTypes = XDocument.Load(entryStream);
            XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
            Assert.NotNull(contentTypes.Root?.Elements(ct + "Override").FirstOrDefault(e => e.Attribute("PartName")?.Value == "/visio/masters/masters.xml" && e.Attribute("ContentType")?.Value == "application/vnd.ms-visio.masters+xml"));
            Assert.NotNull(contentTypes.Root?.Elements(ct + "Override").FirstOrDefault(e => e.Attribute("PartName")?.Value == "/visio/masters/master1.xml" && e.Attribute("ContentType")?.Value == "application/vnd.ms-visio.master+xml"));
        }
    }
}