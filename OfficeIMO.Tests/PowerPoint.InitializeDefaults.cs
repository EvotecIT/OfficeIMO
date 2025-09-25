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

            bool notesMasterExists = false;

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                PresentationPart part = document.PresentationPart!;
                Assert.NotNull(part.Presentation.SlideSize);
                Assert.NotNull(part.Presentation.NotesSize);
                SlideMasterPart master = part.SlideMasterParts.First();
                Assert.NotEmpty(master.SlideLayoutParts);
                ThemePart theme = part.ThemePart!;
                Assert.Same(theme, master.ThemePart);
                Assert.Single(part.Presentation.SlideIdList!.Elements<SlideId>());

                Assert.NotNull(document.ExtendedFilePropertiesPart?.Properties);
                Assert.Equal("Microsoft Office PowerPoint", document.ExtendedFilePropertiesPart?.Properties?.Application?.Text);
                Assert.Equal("1", document.ExtendedFilePropertiesPart?.Properties?.Slides?.Text);
                Assert.Equal("0", document.ExtendedFilePropertiesPart?.Properties?.Notes?.Text);
                Assert.Equal("0", document.ExtendedFilePropertiesPart?.Properties?.HiddenSlides?.Text);
                Assert.NotNull(document.CoreFilePropertiesPart);
                Assert.Equal("OfficeIMO", document.PackageProperties.Creator);
                Assert.Equal("OfficeIMO", document.PackageProperties.LastModifiedBy);
                Assert.NotNull(document.PackageProperties.Created);
                Assert.NotNull(document.PackageProperties.Modified);
                Assert.NotNull(part.PresentationPropertiesPart?.PresentationProperties);
                Assert.True(part.PresentationPropertiesPart!.PresentationProperties!.ShowProperties!.ShowAnimation!.Value);
                Assert.True(part.PresentationPropertiesPart!.PresentationProperties!.ShowProperties!.UseTimings!.Value);
                Assert.False(part.PresentationPropertiesPart!.PresentationProperties!.ShowProperties!.ShowNarration!.Value);
                Assert.NotNull(part.ViewPropertiesPart?.ViewProperties);
                Assert.Equal(15989, part.ViewPropertiesPart!.ViewProperties!.NormalViewProperties!.RestoredLeft!.Size!.Value);
                Assert.False(part.ViewPropertiesPart!.ViewProperties!.NormalViewProperties!.RestoredLeft!.AutoAdjust!.Value);
                Assert.Equal(94660, part.ViewPropertiesPart!.ViewProperties!.NormalViewProperties!.RestoredTop!.Size!.Value);
                Assert.Equal(72008L, part.ViewPropertiesPart!.ViewProperties!.GridSpacing!.Cx!.Value);
                Assert.Equal(72008L, part.ViewPropertiesPart!.ViewProperties!.GridSpacing!.Cy!.Value);
                Assert.Equal("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}", part.TableStylesPart?.TableStyleList?.Default?.Value);
                notesMasterExists = part.NotesMasterPart != null;
            }

            using (FileStream fs = File.OpenRead(filePath))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read)) {
                ZipArchiveEntry entry = zip.GetEntry("[Content_Types].xml")!;
                using Stream s = entry.Open();
                XDocument xml = XDocument.Load(s);
                XNamespace ns = "http://schemas.openxmlformats.org/package/2006/content-types";

                bool HasOverride(string type) => xml.Descendants(ns + "Override")
                    .Any(o => (string?)o.Attribute("ContentType") == type);
                bool HasDefault(string type) => xml.Descendants(ns + "Default")
                    .Any(o => (string?)o.Attribute("ContentType") == type);

                const string pres = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
                bool presentationDefined = HasOverride(pres) ||
                    xml.Descendants(ns + "Default").Any(d => (string?)d.Attribute("ContentType") == pres);
                Assert.True(presentationDefined);
                Assert.True(HasOverride("application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"));
                Assert.True(HasOverride("application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"));
                Assert.True(HasOverride("application/vnd.openxmlformats-officedocument.presentationml.slide+xml"));
                Assert.True(HasOverride("application/vnd.openxmlformats-officedocument.theme+xml"));
                bool notesMasterOverride = HasOverride("application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml");
                Assert.Equal(notesMasterExists, notesMasterOverride);
                Assert.True(HasOverride("application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"));
                Assert.True(HasOverride("application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"));
                Assert.True(HasOverride("application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"));
                Assert.True(HasOverride("application/vnd.openxmlformats-officedocument.extended-properties+xml"));
                Assert.True(HasOverride("application/vnd.openxmlformats-package.core-properties+xml")
                    || HasDefault("application/vnd.openxmlformats-package.core-properties+xml"));
            }

            File.Delete(filePath);
        }
    }
}
