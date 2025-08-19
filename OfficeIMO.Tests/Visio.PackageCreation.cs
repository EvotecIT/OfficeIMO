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
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            using (Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read)) {
                Assert.True(package.PartExists(new Uri("/visio/document.xml", UriKind.Relative)));
                Assert.True(package.PartExists(new Uri("/visio/pages/pages.xml", UriKind.Relative)));
                Assert.True(package.PartExists(new Uri("/visio/pages/page1.xml", UriKind.Relative)));
                Assert.True(package.PartExists(new Uri("/docProps/core.xml", UriKind.Relative)));
                Assert.True(package.PartExists(new Uri("/docProps/app.xml", UriKind.Relative)));
                Assert.True(package.PartExists(new Uri("/docProps/custom.xml", UriKind.Relative)));
                Assert.True(package.PartExists(new Uri("/docProps/thumbnail.emf", UriKind.Relative)));
                Assert.True(package.PartExists(new Uri("/visio/windows.xml", UriKind.Relative)));

                PackageRelationship rel = package.GetRelationshipsByType("http://schemas.microsoft.com/visio/2010/relationships/document").Single();
                Assert.Equal("/visio/document.xml", rel.TargetUri.OriginalString);
                Assert.Equal("rId1", rel.Id);

                PackageRelationship coreRel = package.GetRelationshipsByType("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties").Single();
                Assert.Equal("/docProps/core.xml", coreRel.TargetUri.OriginalString);

                PackageRelationship appRel = package.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties").Single();
                Assert.Equal("/docProps/app.xml", appRel.TargetUri.OriginalString);

                PackageRelationship customRel = package.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties").Single();
                Assert.Equal("/docProps/custom.xml", customRel.TargetUri.OriginalString);

                PackageRelationship thumbRel = package.GetRelationshipsByType("http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail").Single();
                Assert.Equal("/docProps/thumbnail.emf", thumbRel.TargetUri.OriginalString);

                Assert.Empty(package.GetRelationshipsByType("http://schemas.microsoft.com/visio/2010/relationships/windows"));

                PackagePart documentPart = package.GetPart(new Uri("/visio/document.xml", UriKind.Relative));
                PackageRelationship pagesRel = documentPart.GetRelationshipsByType("http://schemas.microsoft.com/visio/2010/relationships/pages").Single();
                Assert.Equal("rId1", pagesRel.Id);
                Uri pagesUri = PackUriHelper.ResolvePartUri(documentPart.Uri, pagesRel.TargetUri);
                Assert.Equal("/visio/pages/pages.xml", pagesUri.OriginalString);

                PackageRelationship windowsRel = documentPart.GetRelationshipsByType("http://schemas.microsoft.com/visio/2010/relationships/windows").Single();
                Uri windowsUri = PackUriHelper.ResolvePartUri(documentPart.Uri, windowsRel.TargetUri);
                Assert.Equal("/visio/windows.xml", windowsUri.OriginalString);

                PackagePart pagesPart = package.GetPart(pagesUri);
                PackageRelationship pageRel = pagesPart.GetRelationshipsByType("http://schemas.microsoft.com/visio/2010/relationships/page").Single();
                Assert.Equal("rId1", pageRel.Id);
                pageUri = PackUriHelper.ResolvePartUri(pagesPart.Uri, pageRel.TargetUri);
                Assert.Equal("/visio/pages/page1.xml", pageUri.OriginalString);

                XDocument pagesDoc = XDocument.Load(pagesPart.GetStream());
                XElement? pageElement = pagesDoc.Root?.Element(ns + "Page");
                Assert.Equal("1", pageElement?.Attribute("ID")?.Value);
                Assert.Null(pageElement?.Attribute("RelId"));
                string? relId = pageElement?.Element(ns + "Rel")?.Attribute(rNs + "id")?.Value;
                Assert.Equal(pageRel.Id, relId);

                PackagePart pagePart = package.GetPart(pageUri);
                using (Stream pageStream = pagePart.GetStream()) {
                    pageDoc = XDocument.Load(pageStream);
                }
            }

            using (FileStream zipStream = File.OpenRead(filePath))
            using (ZipArchive archive = new(zipStream, ZipArchiveMode.Read)) {
                ZipArchiveEntry entry = archive.GetEntry("[Content_Types].xml");
                Assert.NotNull(entry);
                using Stream entryStream = entry.Open();
                XDocument contentTypes = XDocument.Load(entryStream);
                XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
                Assert.NotNull(contentTypes.Root?.Elements(ct + "Default").FirstOrDefault(e => e.Attribute("Extension")?.Value == "xml" && e.Attribute("ContentType")?.Value == "application/xml"));
                Assert.NotNull(contentTypes.Root?.Elements(ct + "Default").FirstOrDefault(e => e.Attribute("Extension")?.Value == "emf" && e.Attribute("ContentType")?.Value == "image/x-emf"));
                Assert.NotNull(contentTypes.Root?.Elements(ct + "Override").FirstOrDefault(e => e.Attribute("PartName")?.Value == "/visio/document.xml" && e.Attribute("ContentType")?.Value == "application/vnd.ms-visio.drawing.main+xml"));
                Assert.NotNull(contentTypes.Root?.Elements(ct + "Override").FirstOrDefault(e => e.Attribute("PartName")?.Value == "/visio/pages/pages.xml" && e.Attribute("ContentType")?.Value == "application/vnd.ms-visio.pages+xml"));
                Assert.NotNull(contentTypes.Root?.Elements(ct + "Override").FirstOrDefault(e => e.Attribute("PartName")?.Value == "/visio/pages/page1.xml" && e.Attribute("ContentType")?.Value == "application/vnd.ms-visio.page+xml"));
                Assert.NotNull(contentTypes.Root?.Elements(ct + "Override").FirstOrDefault(e => e.Attribute("PartName")?.Value == "/docProps/core.xml" && e.Attribute("ContentType")?.Value == "application/vnd.openxmlformats-package.core-properties+xml"));
                Assert.NotNull(contentTypes.Root?.Elements(ct + "Override").FirstOrDefault(e => e.Attribute("PartName")?.Value == "/docProps/app.xml" && e.Attribute("ContentType")?.Value == "application/vnd.openxmlformats-officedocument.extended-properties+xml"));
                Assert.NotNull(contentTypes.Root?.Elements(ct + "Override").FirstOrDefault(e => e.Attribute("PartName")?.Value == "/docProps/custom.xml" && e.Attribute("ContentType")?.Value == "application/vnd.openxmlformats-officedocument.custom-properties+xml"));
                Assert.NotNull(contentTypes.Root?.Elements(ct + "Override").FirstOrDefault(e => e.Attribute("PartName")?.Value == "/docProps/thumbnail.emf" && e.Attribute("ContentType")?.Value == "image/x-emf"));
                Assert.NotNull(contentTypes.Root?.Elements(ct + "Override").FirstOrDefault(e => e.Attribute("PartName")?.Value == "/visio/windows.xml" && e.Attribute("ContentType")?.Value == "application/vnd.ms-visio.windows+xml"));

                XElement shape = pageDoc.Root?.Element(ns + "Shapes")?.Element(ns + "Shape");
                Assert.Equal("1", shape?.Attribute("ID")?.Value);
                Assert.Equal("Rectangle", shape?.Element(ns + "Text")?.Value);
            }

            var issues = VisioValidator.Validate(filePath);
            Assert.Empty(issues);
        }
    }
}

