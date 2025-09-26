using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
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
    }
}
