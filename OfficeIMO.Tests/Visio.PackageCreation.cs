using System;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioPackageCreation {
        [Fact]
        public void CreatesPackageWithExpectedParts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Rectangle"));
            document.Save(filePath);

            XDocument pageDoc;
            Uri pageUri;

            using (Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read)) {
                Assert.True(package.PartExists(new Uri("/visio/document.xml", UriKind.Relative)));
                Assert.True(package.PartExists(new Uri("/visio/pages/pages.xml", UriKind.Relative)));
                Assert.True(package.PartExists(new Uri("/visio/pages/page1.xml", UriKind.Relative)));

                PackageRelationship rel = package.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument").Single();
                Assert.Equal("/visio/document.xml", rel.TargetUri.OriginalString);

                PackagePart documentPart = package.GetPart(new Uri("/visio/document.xml", UriKind.Relative));
                PackageRelationship pagesRel = documentPart.GetRelationshipsByType("http://schemas.microsoft.com/visio/2010/relationships/pages").Single();
                Uri pagesUri = PackUriHelper.ResolvePartUri(documentPart.Uri, pagesRel.TargetUri);
                Assert.Equal("/visio/pages/pages.xml", pagesUri.OriginalString);

                PackagePart pagesPart = package.GetPart(pagesUri);
                PackageRelationship pageRel = pagesPart.GetRelationshipsByType("http://schemas.microsoft.com/visio/2010/relationships/page").Single();
                pageUri = PackUriHelper.ResolvePartUri(pagesPart.Uri, pageRel.TargetUri);
                Assert.Equal("/visio/pages/page1.xml", pageUri.OriginalString);

                PackagePart pagePart = package.GetPart(pageUri);
                pageDoc = XDocument.Load(pagePart.GetStream());
            }

            using FileStream zipStream = File.OpenRead(filePath);
            using ZipArchive archive = new(zipStream, ZipArchiveMode.Read);
            ZipArchiveEntry entry = archive.GetEntry("[Content_Types].xml");
            Assert.NotNull(entry);
            using Stream entryStream = entry.Open();
            XDocument contentTypes = XDocument.Load(entryStream);
            XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
            bool hasDocOverride = contentTypes.Root?.Elements(ct + "Override").Any(e => e.Attribute("PartName")?.Value == "/visio/document.xml" && e.Attribute("ContentType")?.Value == "application/vnd.ms-visio.document.main+xml") == true;
            bool hasDocDefault = contentTypes.Root?.Elements(ct + "Default").Any(e => e.Attribute("Extension")?.Value == "xml" && e.Attribute("ContentType")?.Value == "application/vnd.ms-visio.document.main+xml") == true;
            Assert.True(hasDocOverride || hasDocDefault);
            Assert.NotNull(contentTypes.Root?.Elements(ct + "Override").FirstOrDefault(e => e.Attribute("PartName")?.Value == "/visio/pages/pages.xml"));
            Assert.NotNull(contentTypes.Root?.Elements(ct + "Override").FirstOrDefault(e => e.Attribute("PartName")?.Value == "/visio/pages/page1.xml"));

            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement shape = pageDoc.Root?.Element(ns + "Shapes")?.Element(ns + "Shape");
            Assert.Equal("1", shape?.Attribute("ID")?.Value);
            Assert.Equal("Rectangle", shape?.Element(ns + "Text")?.Value);
        }
    }
}

