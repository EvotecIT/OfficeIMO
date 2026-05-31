using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioSvgExport {
        [Fact]
        public void DocumentCanExportFirstPageToHeadlessSvg() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Diagram").Size(6, 4);
            VisioShape start = page.AddRectangle(1, 2, 1.5, 0.75, "Start");
            start.FillColor = OfficeColor.FromRgb(238, 247, 255);
            start.LineColor = OfficeColor.FromRgb(37, 99, 235);
            start.TextStyle = new VisioTextStyle {
                FontFamily = "Aptos",
                Size = 12,
                Bold = true,
                Color = OfficeColor.FromRgb(17, 24, 39)
            };

            VisioShape decision = page.AddDiamond(4, 2, 1.2, 1.2, "OK?");
            VisioConnector connector = page.AddConnector(start, decision, ConnectorKind.RightAngle, VisioSide.Right, VisioSide.Left);
            connector.EndArrow = EndArrow.Arrow;
            connector.Label = "yes";
            connector.LabelPlacement = VisioConnectorLabelPlacement.Along(0.5, 0, 0.2);

            string svg = document.ToSvg(new VisioSvgSaveOptions {
                PixelsPerInch = 100,
                BackgroundColor = null
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement root = parsed.Root!;
            Assert.Equal("600", root.Attribute("width")!.Value);
            Assert.Equal("400", root.Attribute("height")!.Value);
            Assert.Equal("0 0 600 400", root.Attribute("viewBox")!.Value);
            Assert.Contains(root.Descendants(ns + "g"), g => (string?)g.Attribute("data-visio-shape-id") == start.Id);
            Assert.Contains(root.Descendants(ns + "g"), g => (string?)g.Attribute("data-visio-connector-id") == connector.Id);
            Assert.Contains(root.Descendants(ns + "text"), text => text.Value.IndexOf("Start", StringComparison.Ordinal) >= 0);
            Assert.Contains(root.Descendants(ns + "text"), text => text.Value.IndexOf("yes", StringComparison.Ordinal) >= 0);
            Assert.Contains(root.Descendants(ns + "path"), path => ((string?)path.Attribute("marker-end")) == "url(#officeimo-visio-arrow)");
        }

        [Fact]
        public void PageCanSaveHeadlessSvgToFileAndStream() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Export").Size(2, 1);
            page.AddEllipse(1, 0.5, 1, 0.5, "Node");

            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".svg");
            try {
                page.SaveAsSvg(path, new VisioSvgSaveOptions { IncludeXmlDeclaration = true });
                string fileText = File.ReadAllText(path);
                Assert.StartsWith("<?xml", fileText, StringComparison.Ordinal);
                Assert.True(fileText.IndexOf("<ellipse", StringComparison.Ordinal) >= 0);

                using MemoryStream stream = new();
                document.SaveAsSvg(stream);
                string streamText = Encoding.UTF8.GetString(stream.ToArray());
                Assert.True(streamText.IndexOf("<svg", StringComparison.Ordinal) >= 0);
                Assert.True(streamText.IndexOf("Node", StringComparison.Ordinal) >= 0);
            } finally {
                if (File.Exists(path)) {
                    File.Delete(path);
                }
            }
        }
    }
}
