using System;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioAssetSamples {
        private static string AssetsPath => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets"));

        [Fact]
        public void EmptyDocumentMatchesAsset() {
            string target = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            page.PageWidth = 8.26771653543307;
            page.PageHeight = 11.69291338582677;
            page.ViewCenterX = 4.1233127451916;
            page.ViewCenterY = 5.8492688900245;
            document.Save(target);

            using FileStream expectedStream = File.OpenRead(Path.Combine(AssetsPath, "DrawingEmpty.vsdx"));
            using ZipArchive expected = new(expectedStream, ZipArchiveMode.Read);
            using FileStream actualStream = File.OpenRead(target);
            using ZipArchive actual = new(actualStream, ZipArchiveMode.Read);
            AssertXmlEqual(expected, actual, "visio/pages/pages.xml");
            AssertXmlEqual(expected, actual, "visio/pages/page1.xml");
        }

        [Fact(Skip = "Rectangle output not yet finalized")]
        public void RectangleDocumentMatchesAsset() {
            string target = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            page.PageWidth = 11.69291338582677;
            page.PageHeight = 8.26771653543307;
            page.ViewCenterX = 5.8424184863857;
            page.ViewCenterY = 4.133858091015;
            page.Shapes.Add(new VisioShape("1") {
                NameU = "Rectangle",
                PinX = 2.047244040636296,
                PinY = 6.73228320203895
            });
            document.Save(target);

            using FileStream expectedStream = File.OpenRead(Path.Combine(AssetsPath, "DrawingWithRectangle.vsdx"));
            using ZipArchive expected = new(expectedStream, ZipArchiveMode.Read);
            using FileStream actualStream = File.OpenRead(target);
            using ZipArchive actual = new(actualStream, ZipArchiveMode.Read);
            AssertXmlEqual(expected, actual, "visio/pages/pages.xml");
            AssertXmlEqual(expected, actual, "visio/pages/page1.xml");
        }

        private static void AssertXmlEqual(ZipArchive expected, ZipArchive actual, string entryPath) {
            ZipArchiveEntry? expectedEntry = expected.GetEntry(entryPath);
            ZipArchiveEntry? actualEntry = actual.GetEntry(entryPath);
            Assert.NotNull(expectedEntry);
            Assert.NotNull(actualEntry);
            using Stream expectedStream = expectedEntry!.Open();
            using Stream actualStream = actualEntry!.Open();
            XDocument expectedDoc = XDocument.Load(expectedStream);
            XDocument actualDoc = XDocument.Load(actualStream);
            Assert.True(XNode.DeepEquals(Normalize(expectedDoc.Root!), Normalize(actualDoc.Root!)), $"{entryPath} differed");
        }

        private static XElement Normalize(XElement element) {
            return new XElement(element.Name,
                element.Attributes().OrderBy(a => a.Name.ToString()),
                element.Elements().Select(Normalize));
        }
    }
}

