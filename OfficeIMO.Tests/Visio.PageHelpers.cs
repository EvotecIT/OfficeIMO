using System;
using System.IO;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioPageHelpers {
        [Fact]
        public void AddShapeAndConnectorPopulateCollections() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1", 8.5, 11);
            VisioMaster master = new("1", "Rectangle", new VisioShape("1"));

            VisioShape shape1 = page.AddShape("1", master, 1, 1, 1, 1, "Start");
            VisioShape shape2 = page.AddShape("2", master, 2, 1, 1, 1, "End");

            Assert.Equal(8.5, page.Width);
            Assert.Equal(11, page.Height);
            Assert.Equal(2, page.Shapes.Count);
            Assert.Contains(shape1, page.Shapes);
            Assert.Contains(shape2, page.Shapes);

            VisioConnector connector = page.AddConnector("3", shape1, shape2, ConnectorKind.Straight);

            Assert.Single(page.Connectors);
            Assert.Contains(connector, page.Connectors);
            Assert.Equal(ConnectorKind.Straight, connector.Kind);
            document.Save();
        }

        [Fact]
        public void SizeAndGridHelpersSetProperties() {
            VisioPage page = new("Page-1");
            page.Size(10, 5).Grid(true, false);
            Assert.Equal(10, page.Width);
            Assert.Equal(5, page.Height);
            Assert.True(page.GridVisible);
            Assert.False(page.Snap);
        }
    }
}
