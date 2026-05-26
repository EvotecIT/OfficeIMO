using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Tests {
    public class VisioBackgroundPageTests {
        private static string AssetsPath => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets"));

        [Fact]
        public void BackgroundPagesSaveLoadAndRoundTripBackPageReferences() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string roundTripPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage background = document.AddBackgroundPage("Brand background", 11, 8.5);
            background.AddRectangle(5.5, 8.05, 10.5, 0.45, "OfficeIMO generated")
                .Protect(protection => protection.Size().Position().Text().Selection());
            background.AddRectangle(5.5, 0.35, 10.5, 0.25, string.Empty).FillColor = Color.LightGray;

            VisioPage architecture = document.AddPage("Architecture", 11, 8.5);
            architecture.SetBackgroundPage(background);
            architecture.AddRectangle(3.5, 4.8, 2.2, 1, "API");
            architecture.AddRectangle(7.5, 4.8, 2.2, 1, "Worker");

            VisioPage operations = document.AddPage("Operations", 11, 8.5);
            operations.SetBackgroundPage(background);
            operations.AddRectangle(5.5, 4.8, 2.2, 1, "Runbook");

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertPagesXml(filePath, background.Id, architecture.Id, operations.Id);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedBackground = loaded.Pages.Single(page => page.Name == "Brand background");
            VisioPage loadedArchitecture = loaded.Pages.Single(page => page.Name == "Architecture");
            VisioPage loadedOperations = loaded.Pages.Single(page => page.Name == "Operations");
            Assert.True(loadedBackground.IsBackground);
            Assert.Same(loadedBackground, loadedArchitecture.BackgroundPage);
            Assert.Same(loadedBackground, loadedOperations.BackgroundPage);

            loaded.Save(roundTripPath);

            Assert.Empty(VisioValidator.Validate(roundTripPath));
            AssertPagesXml(roundTripPath, loadedBackground.Id, loadedArchitecture.Id, loadedOperations.Id);
        }

        [Fact]
        public void LoadsBackgroundPagesFromVisioAuthoredFixture() {
            string template = Path.Combine(AssetsPath, "VisioTemplates", "DrawingWithLotsOfShapresAndArrows.vsdx");

            VisioDocument document = VisioDocument.Load(template);

            Assert.True(document.Pages.Count(page => page.IsBackground) >= 3);
            Assert.True(document.Pages.Count(page => page.BackgroundPage != null) >= 3);
            Assert.All(document.Pages.Where(page => page.BackgroundPage != null), page => Assert.True(page.BackgroundPage!.IsBackground));
        }

        [Fact]
        public void BackgroundPageCanReferenceAnotherBackgroundPageOnRoundTrip() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string roundTripPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage baseBackground = document.AddBackgroundPage("Base", 11, 8.5);
            VisioPage overlayBackground = document.AddBackgroundPage("Overlay", 11, 8.5);
            overlayBackground.SetBackgroundPage(baseBackground);
            VisioPage foreground = document.AddPage("Foreground", 11, 8.5);
            foreground.SetBackgroundPage(overlayBackground);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedBase = loaded.Pages.Single(page => page.Name == "Base");
            VisioPage loadedOverlay = loaded.Pages.Single(page => page.Name == "Overlay");
            VisioPage loadedForeground = loaded.Pages.Single(page => page.Name == "Foreground");
            Assert.Same(loadedBase, loadedOverlay.BackgroundPage);
            Assert.Same(loadedOverlay, loadedForeground.BackgroundPage);

            loaded.Save(roundTripPath);
            Assert.Empty(VisioValidator.Validate(roundTripPath));
            AssertChainedBackgroundXml(roundTripPath, loadedBase.Id, loadedOverlay.Id, loadedForeground.Id);
        }

        [Fact]
        public void BackgroundPageMustBelongToSameDocument() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage foreground = document.AddPage("Foreground", 11, 8.5);
            VisioDocument otherDocument = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage foreignBackground = otherDocument.AddBackgroundPage("Foreign", 11, 8.5);
            VisioPage detachedBackground = new("Detached", 11, 8.5);

            Assert.Throws<InvalidOperationException>(() => foreground.SetBackgroundPage(foreignBackground));
            Assert.Throws<InvalidOperationException>(() => foreground.SetBackgroundPage(detachedBackground));
        }

        private static void AssertPagesXml(string filePath, int backgroundId, int architectureId, int operationsId) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument pages = ReadXml(archive, "visio/pages/pages.xml");

            XElement background = PageById(pages, ns, backgroundId);
            Assert.Equal("1", background.Attribute("Background")?.Value);
            Assert.Null(background.Attribute("BackPage"));

            XElement architecture = PageById(pages, ns, architectureId);
            Assert.Equal(backgroundId.ToString(), architecture.Attribute("BackPage")?.Value);
            Assert.Null(architecture.Attribute("Background"));

            XElement operations = PageById(pages, ns, operationsId);
            Assert.Equal(backgroundId.ToString(), operations.Attribute("BackPage")?.Value);
            Assert.Null(operations.Attribute("Background"));
        }

        private static void AssertChainedBackgroundXml(string filePath, int baseBackgroundId, int overlayBackgroundId, int foregroundId) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument pages = ReadXml(archive, "visio/pages/pages.xml");

            XElement baseBackground = PageById(pages, ns, baseBackgroundId);
            Assert.Equal("1", baseBackground.Attribute("Background")?.Value);
            Assert.Null(baseBackground.Attribute("BackPage"));

            XElement overlayBackground = PageById(pages, ns, overlayBackgroundId);
            Assert.Equal("1", overlayBackground.Attribute("Background")?.Value);
            Assert.Equal(baseBackgroundId.ToString(), overlayBackground.Attribute("BackPage")?.Value);

            XElement foreground = PageById(pages, ns, foregroundId);
            Assert.Equal(overlayBackgroundId.ToString(), foreground.Attribute("BackPage")?.Value);
            Assert.Null(foreground.Attribute("Background"));
        }

        private static XElement PageById(XDocument pages, XNamespace ns, int id) {
            return pages.Root!.Elements(ns + "Page")
                .Single(page => (string?)page.Attribute("ID") == id.ToString());
        }

        private static XDocument ReadXml(ZipArchive archive, string entryName) {
            ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException("Missing " + entryName);
            using Stream stream = entry.Open();
            return XDocument.Load(stream);
        }
    }
}
