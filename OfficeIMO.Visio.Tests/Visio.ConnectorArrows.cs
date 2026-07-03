using System;
using System.IO;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioConnectorArrows {
        [Fact]
        public void ConnectorRoundTripsKindArrowsAndLabel() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");
            VisioShape start = new("1", 1, 1, 1, 1, "Start");
            VisioShape end = new("2", 3, 2, 1, 1, "End");
            page.Shapes.Add(start);
            page.Shapes.Add(end);

            VisioConnector connector = new VisioConnector(start, end) {
                Kind = ConnectorKind.RightAngle,
                BeginArrow = EndArrow.Arrow,
                EndArrow = EndArrow.Triangle,
                Label = "A to B"
            };
            page.Connectors.Add(connector);
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector loadedConnector = loaded.Pages[0].Connectors[0];

            Assert.Equal(ConnectorKind.RightAngle, loadedConnector.Kind);
            Assert.Equal(EndArrow.Arrow, loadedConnector.BeginArrow);
            Assert.Equal(EndArrow.Triangle, loadedConnector.EndArrow);
            Assert.Equal("A to B", loadedConnector.Label);
        }
    }
}
