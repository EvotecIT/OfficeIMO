using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioPageScaleTests {
        [Fact]
        public void CustomScalesPersistAcrossSaveAndLoad() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1", 8.5, 11, VisioMeasurementUnit.Inches);
            page.ScaleMeasurementUnit = VisioMeasurementUnit.Centimeters;
            page.PageScale = new VisioScaleSetting(2.5, VisioMeasurementUnit.Centimeters);
            page.DrawingScale = new VisioScaleSetting(5, VisioMeasurementUnit.Inches);
            document.Save();

            VisioDocument reloaded = VisioDocument.Load(filePath);
            VisioPage reloadedPage = reloaded.Pages[0];

            Assert.Equal(VisioMeasurementUnit.Centimeters, reloadedPage.ScaleMeasurementUnit);
            Assert.Equal(VisioMeasurementUnit.Centimeters, reloadedPage.PageScale.Unit);
            Assert.Equal(2.5, reloadedPage.PageScale.Value, 5);
            Assert.Equal(VisioMeasurementUnit.Inches, reloadedPage.DrawingScale.Unit);
            Assert.Equal(5, reloadedPage.DrawingScale.Value, 5);

            using FileStream stream = File.OpenRead(filePath);
            using ZipArchive archive = new(stream, ZipArchiveMode.Read);
            ZipArchiveEntry? pagesEntry = archive.GetEntry("visio/pages/pages.xml");
            Assert.NotNull(pagesEntry);
            using Stream entryStream = pagesEntry!.Open();
            XDocument pagesXml = XDocument.Load(entryStream);
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            var cells = pagesXml.Root!
                .Element(v + "Page")!
                .Element(v + "PageSheet")!
                .Elements(v + "Cell")
                .ToDictionary(e => (string)e.Attribute("N")!, e => (val: (string?)e.Attribute("V"), unit: (string?)e.Attribute("U")));

            string expectedPageScale = XmlConvert.ToString(2.5d.ToInches(VisioMeasurementUnit.Centimeters));
            Assert.Equal("CM", cells["PageScale"].unit);
            Assert.Equal(expectedPageScale, cells["PageScale"].val);

            string expectedDrawingScale = XmlConvert.ToString(5d.ToInches(VisioMeasurementUnit.Inches));
            Assert.Equal("IN", cells["DrawingScale"].unit);
            Assert.Equal(expectedDrawingScale, cells["DrawingScale"].val);
        }

        [Fact]
        public void MetricPagesWithoutExplicitScaleUnitsReloadCorrectly() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Metric", 20, 30, VisioMeasurementUnit.Millimeters);
            page.ScaleMeasurementUnit = VisioMeasurementUnit.Millimeters;
            page.PageScale = VisioScaleSetting.FromUnit(VisioMeasurementUnit.Millimeters);
            page.DrawingScale = VisioScaleSetting.FromUnit(VisioMeasurementUnit.Millimeters);
            document.Save();

            using (ZipArchive archive = ZipFile.Open(filePath, ZipArchiveMode.Update)) {
                ZipArchiveEntry? pagesEntry = archive.GetEntry("visio/pages/pages.xml");
                Assert.NotNull(pagesEntry);

                XDocument pagesXml;
                XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
                using (Stream entryStream = pagesEntry!.Open()) {
                    pagesXml = XDocument.Load(entryStream);
                }

                var cellQuery = pagesXml.Root!
                    .Element(v + "Page")!
                    .Element(v + "PageSheet")!
                    .Elements(v + "Cell")
                    .ToList();

                foreach (XElement cell in cellQuery) {
                    string? name = cell.Attribute("N")?.Value;
                    if (name == "PageScale" || name == "DrawingScale") {
                        cell.Attribute("U")?.Remove();
                    } else if (name == "PageWidth" || name == "PageHeight") {
                        cell.SetAttributeValue("U", "MM");
                    }
                }

                Assert.Equal("MM", cellQuery.First(c => (string?)c.Attribute("N") == "PageWidth").Attribute("U")?.Value);
                Assert.Equal("MM", cellQuery.First(c => (string?)c.Attribute("N") == "PageHeight").Attribute("U")?.Value);
                Assert.Contains("U=\"MM\"", pagesXml.ToString(SaveOptions.DisableFormatting));

                pagesEntry.Delete();
                ZipArchiveEntry replacement = archive.CreateEntry("visio/pages/pages.xml");
                using Stream replacementStream = replacement.Open();
                using StreamWriter writer = new(replacementStream, new UTF8Encoding(false));
                if (pagesXml.Declaration != null) {
                    writer.Write(pagesXml.Declaration);
                }
                writer.Write(pagesXml.ToString(SaveOptions.DisableFormatting));
            }

            VisioDocument reloaded = VisioDocument.Load(filePath);
            VisioPage reloadedPage = reloaded.Pages[0];

            Assert.Equal(VisioMeasurementUnit.Millimeters, reloadedPage.ScaleMeasurementUnit);
            Assert.Equal(VisioMeasurementUnit.Millimeters, reloadedPage.PageScale.Unit);
            Assert.Equal(1, reloadedPage.PageScale.Value, 5);
            Assert.Equal(VisioMeasurementUnit.Millimeters, reloadedPage.DrawingScale.Unit);
            Assert.Equal(1, reloadedPage.DrawingScale.Value, 5);
        }

        [Fact]
        public void ChangingScaleMeasurementUnitConvertsMatchingOverrides() {
            VisioPage page = new("Demo");
            page.ScaleMeasurementUnit = VisioMeasurementUnit.Centimeters;
            page.PageScale = new VisioScaleSetting(2.5, VisioMeasurementUnit.Centimeters);
            page.DrawingScale = new VisioScaleSetting(7.5, VisioMeasurementUnit.Centimeters);

            page.ScaleMeasurementUnit = VisioMeasurementUnit.Millimeters;

            Assert.Equal(VisioMeasurementUnit.Millimeters, page.PageScale.Unit);
            Assert.Equal(25, page.PageScale.Value, 5);
            Assert.Equal(VisioMeasurementUnit.Millimeters, page.DrawingScale.Unit);
            Assert.Equal(75, page.DrawingScale.Value, 5);
        }

        [Fact]
        public void ChangingScaleMeasurementUnitLeavesOtherUnitsUntouched() {
            VisioPage page = new("Mixed");
            page.ScaleMeasurementUnit = VisioMeasurementUnit.Centimeters;
            page.PageScale = new VisioScaleSetting(3, VisioMeasurementUnit.Inches);

            page.ScaleMeasurementUnit = VisioMeasurementUnit.Millimeters;

            Assert.Equal(VisioMeasurementUnit.Inches, page.PageScale.Unit);
            Assert.Equal(3, page.PageScale.Value, 5);
        }

        [Fact]
        public void SettingInvalidScaleMeasurementUnitThrows() {
            VisioPage page = new("Invalid");
            Assert.Throws<ArgumentOutOfRangeException>(() => page.ScaleMeasurementUnit = (VisioMeasurementUnit)42);
        }
    }
}
