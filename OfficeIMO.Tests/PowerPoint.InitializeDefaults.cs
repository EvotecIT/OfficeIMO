using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointInitializeDefaults {
        [Fact]
        public void DefaultPartsAndContentTypesAreCreated() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AddSlide();
                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                PresentationPart part = document.PresentationPart!;
                Assert.NotNull(part.Presentation.SlideSize);
                Assert.NotNull(part.Presentation.NotesSize);
                SlideMasterPart master = part.SlideMasterParts.First();
                Assert.NotEmpty(master.SlideLayoutParts);
                ThemePart theme = part.ThemePart!;
                Assert.Same(theme, master.ThemePart);
                Assert.Single(part.Presentation.SlideIdList!.Elements<SlideId>());
            }

            using (FileStream fs = File.OpenRead(filePath))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read)) {
                ZipArchiveEntry entry = zip.GetEntry("[Content_Types].xml")!;
                using Stream s = entry.Open();
                XDocument xml = XDocument.Load(s);
                XNamespace ns = "http://schemas.openxmlformats.org/package/2006/content-types";

                bool HasOverride(string type) => xml.Descendants(ns + "Override")
                    .Any(o => (string?)o.Attribute("ContentType") == type);

                const string pres = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
                bool presentationDefined = HasOverride(pres) ||
                    xml.Descendants(ns + "Default").Any(d => (string?)d.Attribute("ContentType") == pres);
                Assert.True(presentationDefined);
                Assert.True(HasOverride("application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"));
                Assert.True(HasOverride("application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"));
                Assert.True(HasOverride("application/vnd.openxmlformats-officedocument.presentationml.slide+xml"));
                Assert.True(HasOverride("application/vnd.openxmlformats-officedocument.theme+xml"));
                Assert.True(HasOverride("application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"));
                Assert.True(HasOverride("application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"));
            }

            File.Delete(filePath);
        }
    }
}
