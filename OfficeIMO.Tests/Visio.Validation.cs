using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioValidation {
        [Fact]
        public void SavedDocumentValidatesAndOpens() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Start"));
            document.Save();

            var issues = VisioValidator.Validate(filePath);
            Assert.Empty(issues);

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Single(loaded.Pages);
        }

        [Fact]
        public void MultiPageDocumentValidatesAndLoads() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page1 = document.AddPage("Page-1");
            page1.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Rectangle"));

            VisioPage page2 = document.AddPage("Page-2");
            page2.Shapes.Add(new VisioShape("2", 2, 2, 1.5, 1, "Ellipse"));
            document.Save();

            var issues = VisioValidator.Validate(filePath);
            Assert.Empty(issues);

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(2, loaded.Pages.Count);
        }

        [Fact]
        public void ValidatorDetectsMissingPagePartInMultiPagePackage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage first = document.AddPage("Page-1");
            first.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Start"));
            VisioPage second = document.AddPage("Page-2");
            second.Shapes.Add(new VisioShape("2", 2, 2, 1, 1, "End"));
            document.Save();

            using (FileStream stream = new(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            using (ZipArchive archive = new(stream, ZipArchiveMode.Update)) {
                ZipArchiveEntry? page2Entry = archive.GetEntry("visio/pages/page2.xml");
                Assert.NotNull(page2Entry);
                page2Entry!.Delete();
            }

            var issues = VisioValidator.Validate(filePath);
            Assert.Contains(issues, issue => issue.Contains("Page ID 1 (Page-2)") && issue.Contains("page2.xml") && issue.Contains("missing", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void ValidatorDetectsMissingPageOverrideInMultiPagePackage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page1 = document.AddPage("Page-1");
            page1.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Start"));
            VisioPage page2 = document.AddPage("Page-2");
            page2.Shapes.Add(new VisioShape("2", 2, 2, 1, 1, "End"));
            document.Save();

            using (FileStream stream = new(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            using (ZipArchive archive = new(stream, ZipArchiveMode.Update)) {
                ZipArchiveEntry? entry = archive.GetEntry("[Content_Types].xml");
                Assert.NotNull(entry);
                XDocument contentTypes;
                using (Stream entryStream = entry!.Open()) {
                    contentTypes = XDocument.Load(entryStream);
                }

                entry.Delete();
                XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
                contentTypes.Root!
                    .Elements(ct + "Override")
                    .Where(e => string.Equals((string?)e.Attribute("PartName"), "/visio/pages/page2.xml", StringComparison.OrdinalIgnoreCase))
                    .Remove();

                ZipArchiveEntry newEntry = archive.CreateEntry("[Content_Types].xml");
                using Stream newEntryStream = newEntry.Open();
                contentTypes.Save(newEntryStream);
            }

            var issues = VisioValidator.Validate(filePath);
            Assert.Contains(issues, issue => issue.Contains("Page ID 1 (Page-2)") && issue.Contains("Missing Override", StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(issues, issue => issue.Contains("Page ID 0 (Page-1)"));
        }
    }
}

