using System;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioThemeTests {
        [Fact]
        public void AddsThemePart() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = new();
            document.Theme = new VisioTheme { Name = "Office Theme" };
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, string.Empty));
            document.Save(filePath);

            using (Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read)) {
                Assert.True(package.PartExists(new Uri("/visio/theme/theme1.xml", UriKind.Relative)));

                PackagePart documentPart = package.GetPart(new Uri("/visio/document.xml", UriKind.Relative));
                PackageRelationship themeRel = documentPart.GetRelationshipsByType("http://schemas.microsoft.com/visio/2010/relationships/theme").Single();
                Uri themeUri = PackUriHelper.ResolvePartUri(documentPart.Uri, themeRel.TargetUri);
                Assert.Equal("/visio/theme/theme1.xml", themeUri.OriginalString);
            }

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal("Office Theme", loaded.Theme?.Name);

            using FileStream zipStream = File.OpenRead(filePath);
            using ZipArchive archive = new(zipStream, ZipArchiveMode.Read);
            ZipArchiveEntry entry = archive.GetEntry("[Content_Types].xml")!;
            using Stream entryStream = entry.Open();
            XDocument contentTypes = XDocument.Load(entryStream);
            XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
            Assert.NotNull(contentTypes.Root?.Elements(ct + "Override").FirstOrDefault(e => e.Attribute("PartName")?.Value == "/visio/theme/theme1.xml" && e.Attribute("ContentType")?.Value == "application/vnd.ms-visio.theme+xml"));
        }

        [Fact]
        public void DoesNotAddThemeWhenAbsent() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, string.Empty));
            document.Save(filePath);

            using (Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read)) {
                Assert.False(package.PartExists(new Uri("/visio/theme/theme1.xml", UriKind.Relative)));

                PackagePart documentPart = package.GetPart(new Uri("/visio/document.xml", UriKind.Relative));
                Assert.Empty(documentPart.GetRelationshipsByType("http://schemas.microsoft.com/visio/2010/relationships/theme"));
            }

            using FileStream zipStream = File.OpenRead(filePath);
            using ZipArchive archive = new(zipStream, ZipArchiveMode.Read);
            ZipArchiveEntry entry = archive.GetEntry("[Content_Types].xml")!;
            using Stream entryStream = entry.Open();
            XDocument contentTypes = XDocument.Load(entryStream);
            XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
            Assert.Null(contentTypes.Root?.Elements(ct + "Override").FirstOrDefault(e => e.Attribute("PartName")?.Value == "/visio/theme/theme1.xml"));
        }
    }
}
