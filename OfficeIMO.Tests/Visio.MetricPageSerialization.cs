using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioMetricPageSerializationTests {
        [Fact]
        public void MetricPagesSerializeMillimeterValues() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage metric = document.AddPage("Metric", 210, 297, VisioMeasurementUnit.Millimeters);
            metric.ScaleMeasurementUnit = VisioMeasurementUnit.Millimeters;
            document.Save();

            XDocument pagesXml;
            using (ZipArchive archive = ZipFile.OpenRead(filePath)) {
                ZipArchiveEntry? pagesEntry = archive.GetEntry("visio/pages/pages.xml");
                Assert.NotNull(pagesEntry);

                using Stream entryStream = pagesEntry!.Open();
                pagesXml = XDocument.Load(entryStream);
            }

            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            var cells = pagesXml.Root!
                .Element(v + "Page")!
                .Element(v + "PageSheet")!
                .Elements(v + "Cell")
                .ToDictionary(
                    cell => cell.Attribute("N")!.Value,
                    cell => (
                        value: cell.Attribute("V")?.Value ?? string.Empty,
                        unit: cell.Attribute("U")?.Value
                    ));

            Assert.Equal(("210", "MM"), (cells["PageWidth"].value, cells["PageWidth"].unit));
            Assert.Equal(("297", "MM"), (cells["PageHeight"].value, cells["PageHeight"].unit));
            Assert.Equal(("3", "MM"), (cells["ShdwOffsetX"].value, cells["ShdwOffsetX"].unit));
            Assert.Equal(("-3", "MM"), (cells["ShdwOffsetY"].value, cells["ShdwOffsetY"].unit));
            Assert.Equal(("6.35", "MM"), (cells["PageLeftMargin"].value, cells["PageLeftMargin"].unit));
            Assert.Equal(("6.35", "MM"), (cells["PageRightMargin"].value, cells["PageRightMargin"].unit));
            Assert.Equal(("6.35", "MM"), (cells["PageTopMargin"].value, cells["PageTopMargin"].unit));
            Assert.Equal(("6.35", "MM"), (cells["PageBottomMargin"].value, cells["PageBottomMargin"].unit));

            VisioDocument reloaded = VisioDocument.Load(filePath);
            VisioPage reloadedPage = Assert.Single(reloaded.Pages);

            Assert.Equal(VisioMeasurementUnit.Millimeters, reloadedPage.DefaultUnit);
            Assert.Equal(VisioMeasurementUnit.Millimeters, reloadedPage.ScaleMeasurementUnit);
            Assert.Equal(210, reloadedPage.Width * 25.4, 5);
            Assert.Equal(297, reloadedPage.Height * 25.4, 5);
        }
    }
}
