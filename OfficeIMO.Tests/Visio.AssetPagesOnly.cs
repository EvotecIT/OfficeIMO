using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioAssetPagesOnly {
        private static string AssetsPath => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets"));

        private static XDocument LoadEntry(ZipArchive zip, string entryPath) {
            var e = zip.GetEntry(entryPath);
            Assert.NotNull(e);
            using var s = e!.Open();
            return XDocument.Load(s);
        }

        [Fact]
        public void PagesXmlMatches_DrawingWithShapes() {
            string target = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(target);
            VisioPage page = document.AddPage("Page-1", 29.7, 21, VisioMeasurementUnit.Centimeters);
            page.ViewCenterX = 5.8424184863857;
            page.ViewCenterY = 4.133858091015;
            document.Save();

            using FileStream expectedStream = File.OpenRead(Path.Combine(AssetsPath, "VisioTemplates", "DrawingWithShapes.vsdx"));
            using ZipArchive expected = new(expectedStream, ZipArchiveMode.Read);
            using FileStream actualStream = File.OpenRead(target);
            using ZipArchive actual = new(actualStream, ZipArchiveMode.Read);
            var exp = LoadEntry(expected, "visio/pages/pages.xml");
            var act = LoadEntry(actual, "visio/pages/pages.xml");
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            var ePage = exp.Root!.Element(v + "Page")!;
            var aPage = act.Root!.Element(v + "Page")!;
            string? expectedViewScaleAttr = ePage.Attribute("ViewScale")?.Value;
            string? actualViewScaleAttr = aPage.Attribute("ViewScale")?.Value;
            double expectedViewScale = expectedViewScaleAttr != null ? XmlConvert.ToDouble(expectedViewScaleAttr) : 1;
            double actualViewScale = actualViewScaleAttr != null ? XmlConvert.ToDouble(actualViewScaleAttr) : 1;
            if (double.IsNaN(expectedViewScale) || double.IsInfinity(expectedViewScale) || expectedViewScale <= 0) {
                expectedViewScale = 1;
            }
            Assert.Equal(expectedViewScale, actualViewScale);
            Assert.Equal((string?)ePage.Attribute("ViewCenterX"), (string?)aPage.Attribute("ViewCenterX"));
            Assert.Equal((string?)ePage.Attribute("ViewCenterY"), (string?)aPage.Attribute("ViewCenterY"));
            var eCells = ePage.Element(v + "PageSheet")!.Elements(v + "Cell").ToDictionary(c => (string)c.Attribute("N")!, c => (val: (string?)c.Attribute("V"), unit: (string?)c.Attribute("U")));
            var aCells = aPage.Element(v + "PageSheet")!.Elements(v + "Cell").ToDictionary(c => (string)c.Attribute("N")!, c => (val: (string?)c.Attribute("V"), unit: (string?)c.Attribute("U")));
            // Assert key cells match exactly
            void Eq(string n) { Assert.Equal(eCells[n], aCells[n]); }
            Eq("PageWidth"); Eq("PageHeight"); Eq("ShdwOffsetX"); Eq("ShdwOffsetY");
            Eq("PageScale"); Eq("DrawingScale"); Eq("DrawingSizeType"); Eq("DrawingScaleType");
            Eq("InhibitSnap"); Eq("ShdwType"); Eq("ShdwObliqueAngle"); Eq("ShdwScaleFactor");
            Eq("DrawingResizeType"); Eq("PageShapeSplit"); Eq("ColorSchemeIndex"); Eq("EffectSchemeIndex");
            Eq("ConnectorSchemeIndex"); Eq("FontSchemeIndex"); Eq("ThemeIndex");
            Eq("PageLeftMargin"); Eq("PageRightMargin"); Eq("PageTopMargin"); Eq("PageBottomMargin"); Eq("PrintPageOrientation");
        }

        [Fact]
        public void PagesXmlMatches_DrawingWithInfoAndShapes() {
            string target = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(target);
            VisioPage page = document.AddPage("Page-1"); // default A4 portrait in inches
            page.ViewScale = 0.98777777777778;
            page.ViewCenterX = 4.1233127451916;
            page.ViewCenterY = 5.3993251292534;
            document.Save();

            using FileStream expectedStream = File.OpenRead(Path.Combine(AssetsPath, "VisioTemplates", "DrawingWithInfoAndShapes.vsdx"));
            using ZipArchive expected = new(expectedStream, ZipArchiveMode.Read);
            using FileStream actualStream = File.OpenRead(target);
            using ZipArchive actual = new(actualStream, ZipArchiveMode.Read);
            var exp = LoadEntry(expected, "visio/pages/pages.xml");
            var act = LoadEntry(actual, "visio/pages/pages.xml");
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            var ePage = exp.Root!.Element(v + "Page")!;
            var aPage = act.Root!.Element(v + "Page")!;
            string? expectedViewScaleAttr = ePage.Attribute("ViewScale")?.Value;
            string? actualViewScaleAttr = aPage.Attribute("ViewScale")?.Value;
            Assert.NotNull(expectedViewScaleAttr);
            Assert.NotNull(actualViewScaleAttr);
            Assert.Equal(XmlConvert.ToDouble(expectedViewScaleAttr), XmlConvert.ToDouble(actualViewScaleAttr));
            Assert.Equal((string?)ePage.Attribute("ViewCenterX"), (string?)aPage.Attribute("ViewCenterX"));
            Assert.Equal((string?)ePage.Attribute("ViewCenterY"), (string?)aPage.Attribute("ViewCenterY"));
            var eCells = ePage.Element(v + "PageSheet")!.Elements(v + "Cell").ToDictionary(c => (string)c.Attribute("N")!, c => (val: (string?)c.Attribute("V"), unit: (string?)c.Attribute("U")));
            var aCells = aPage.Element(v + "PageSheet")!.Elements(v + "Cell").ToDictionary(c => (string)c.Attribute("N")!, c => (val: (string?)c.Attribute("V"), unit: (string?)c.Attribute("U")));
            // Check a key subset only for portrait defaults
            void Eq(string n) { Assert.Equal(eCells[n], aCells[n]); }
            Eq("PageWidth"); Eq("PageHeight"); Eq("PageScale"); Eq("DrawingScale");
        }
    }
}
