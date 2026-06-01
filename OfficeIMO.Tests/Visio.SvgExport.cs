using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
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
            XElement startText = root.Descendants(ns + "text").Single(text => text.Value.IndexOf("Start", StringComparison.Ordinal) >= 0);
            Assert.Equal("16.667", startText.Attribute("font-size")!.Value);
            Assert.Contains(root.Descendants(ns + "path"), path => ((string?)path.Attribute("data-officeimo-connector-arrow")) == "end");
        }

        [Fact]
        public void SvgRendererDrawsConnectorArrowheadsWithLineColorOpacity() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Arrow Opacity").Size(4, 2);
            VisioShape source = page.AddRectangle(0.8, 1, 0.4, 0.4, string.Empty);
            VisioShape target = page.AddRectangle(3.2, 1, 0.4, 0.4, string.Empty);
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            connector.BeginArrow = EndArrow.Arrow;
            connector.EndArrow = EndArrow.Arrow;
            connector.LineColor = OfficeColor.FromRgba(37, 99, 235, 128);
            connector.LineWeight = 0.03D;

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement connectorGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-connector-id") == connector.Id);
            XElement connectorPath = connectorGroup.Elements(ns + "path")
                .Single(path => path.Attribute("data-officeimo-connector-arrow") == null);
            Assert.Null(connectorPath.Attribute("marker-start"));
            Assert.Null(connectorPath.Attribute("marker-end"));
            Assert.Equal("#2563EB", connectorPath.Attribute("stroke")!.Value);
            Assert.Equal("0.502", connectorPath.Attribute("stroke-opacity")!.Value);

            XElement[] arrows = connectorGroup.Elements(ns + "path")
                .Where(path => path.Attribute("data-officeimo-connector-arrow") != null)
                .ToArray();
            Assert.Equal(2, arrows.Length);
            Assert.Contains(arrows, arrow => (string?)arrow.Attribute("data-officeimo-connector-arrow") == "start");
            Assert.Contains(arrows, arrow => (string?)arrow.Attribute("data-officeimo-connector-arrow") == "end");
            Assert.All(arrows, arrow => {
                Assert.Equal("#2563EB", arrow.Attribute("fill")!.Value);
                Assert.Equal("0.502", arrow.Attribute("fill-opacity")!.Value);
                Assert.Equal("none", arrow.Attribute("stroke")!.Value);
            });
        }

        [Fact]
        public void SvgRendererSuppressesArrowheadsWhenConnectorLineIsHidden() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Hidden Arrow").Size(4, 2);
            VisioShape source = page.AddRectangle(0.8, 1, 0.4, 0.4, string.Empty);
            VisioShape target = page.AddRectangle(3.2, 1, 0.4, 0.4, string.Empty);
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            connector.BeginArrow = EndArrow.Arrow;
            connector.EndArrow = EndArrow.Arrow;
            connector.LineColor = OfficeColor.FromRgb(220, 38, 38);
            connector.LinePattern = 0;

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement connectorGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-connector-id") == connector.Id);
            XElement connectorPath = connectorGroup.Elements(ns + "path").Single();
            Assert.Equal("none", connectorPath.Attribute("stroke")!.Value);
            Assert.DoesNotContain(connectorGroup.Elements(ns + "path"), path => path.Attribute("data-officeimo-connector-arrow") != null);
        }

        [Fact]
        public void SvgFallbackConnectorsUseVerticalEndpoints() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Vertical").Size(4, 6);
            VisioShape top = page.AddRectangle(2, 4.5, 1, 1, "Top");
            VisioShape bottom = page.AddRectangle(2, 1.5, 1, 1, "Bottom");

            VisioConnector connector = page.AddConnector(top, bottom, ConnectorKind.Straight);
            string svg = page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 100 });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement connectorGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-connector-id") == connector.Id);
            XElement path = connectorGroup.Element(ns + "path")!;
            Assert.Equal("M 200 200 L 200 400", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererPreservesDashedShapeOutlines() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Dashed Shape").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillPattern = 0;
            shape.LineColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LinePattern = 2;

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path").Single();
            Assert.Equal("6 4", path.Attribute("stroke-dasharray")!.Value);
            Assert.NotEqual("none", path.Attribute("stroke")!.Value);
        }

        [Fact]
        public void SvgRendererSuppressesZeroWeightShapeOutlines() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Zero Line").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillPattern = 0;
            shape.LineColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineWeight = 0D;

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path").Single();
            Assert.Equal("none", path.Attribute("fill")!.Value);
            Assert.Equal("none", path.Attribute("stroke")!.Value);
        }

        [Fact]
        public void SvgRendererKeepsPlainFlowchartDataAsParallelogram() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Flowchart Data").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.NameU = "Data";

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path").Single();
            Assert.Equal("M 125 150 L 200 150 L 175 50 L 100 50 Z", path.Attribute("d")!.Value);
            Assert.Null(path.Attribute("data-officeimo-database-geometry"));
        }

        [Fact]
        public void SvgRendererDrawsSemanticDatabaseShapesAsCylinders() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Database Shape").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.NameU = "Data";
            shape.SetUserCell("OfficeIMO.StencilId", "architecture.database", "STR");

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement body = shapeGroup.Elements(ns + "path")
                .Single(path => (string?)path.Attribute("data-officeimo-database-geometry") == "true");
            XElement seam = shapeGroup.Elements(ns + "path")
                .Single(path => (string?)path.Attribute("data-officeimo-database-seam") == "true");
            Assert.Contains(" C ", body.Attribute("d")!.Value);
            Assert.Contains(" C ", seam.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererDrawsChevronShapesAsChevronPolygons() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Chevron Shape").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.NameU = "Chevron";

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path").Single();
            Assert.Equal("M 100 150 L 172 150 L 200 100 L 172 50 L 100 50 L 128 100 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererDrawsStartEndShapesAsTerminatorCapsules() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Terminator Shape").Size(4, 2);
            VisioShape shape = page.AddRectangle(2, 1, 1.6, 0.8, string.Empty);
            shape.NameU = "Start/End";

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            Assert.Empty(shapeGroup.Elements(ns + "ellipse"));
            XElement path = shapeGroup.Elements(ns + "path").Single();
            Assert.StartsWith("M 160 140 L 240 140 L", path.Attribute("d")!.Value, StringComparison.Ordinal);
            Assert.Contains("L 240 60 L 160 60 L", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererDrawsDocumentStencilsAsWavyDocuments() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Document Shape").Size(3, 2);
            VisioShape shape = page.AddStencilShape(VisioStencils.CollaborationBusiness, "collab.document", "doc", 1.5, 1, 1.4, 1, string.Empty);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path").Single();
            string d = path.Attribute("d")!.Value;
            Assert.StartsWith("M 80 50 L 220 50 L 220 136 L", d, StringComparison.Ordinal);
            Assert.Contains("L 161.2 128.3 L", d);
            Assert.DoesNotContain("L 185 50 L 80 50 Z", d, StringComparison.Ordinal);
        }

        [Fact]
        public void SvgRendererDrawsDelayShapesAsDShapes() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Delay Shape").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1.4, 1, string.Empty);
            shape.NameU = "Delay";

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path").Single();
            string d = path.Attribute("d")!.Value;
            Assert.StartsWith("M 80 150 L 170 150 L", d, StringComparison.Ordinal);
            Assert.Contains("L 220 100 L", d);
            Assert.DoesNotContain("L 220 50 L 80 50 Z", d, StringComparison.Ordinal);
        }

        [Fact]
        public void SvgRendererDrawsManualInputShapesAsSlantedQuadrilaterals() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Manual Input Shape").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.NameU = "Manual Input";

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path").Single();
            string d = path.Attribute("d")!.Value;
            Assert.Equal("M 100 150 L 200 150 L 200 75 L 100 50 Z", d);
            Assert.DoesNotContain("L 200 50 L 100 50 Z", d, StringComparison.Ordinal);
        }

        [Fact]
        public void SvgRendererWrapsBoundedTextAndDrawsConnectorLabelBackgrounds() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Labels").Size(5, 3);
            VisioShape source = page.AddRectangle(1.2, 1.5, 1.1, 0.55, "Customer onboarding approval");
            source.TextStyle = new VisioTextStyle {
                Size = 12,
                TextWidth = 0.85,
                TextHeight = 0.42,
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };
            VisioShape target = page.AddRectangle(3.8, 1.5, 1.1, 0.55, "Done");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.RightAngle, VisioSide.Right, VisioSide.Left);
            connector.Label = "manual review required";
            connector.LabelPlacement = VisioConnectorLabelPlacement.Along(0.5, 0, 0.18);
            connector.TextStyle = new VisioTextStyle {
                Size = 9,
                TextWidth = 0.72,
                TextHeight = 0.32
            };

            string svg = page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 100 });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement sourceGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == source.Id);
            XElement shapeText = sourceGroup.Descendants(ns + "text").Single();
            Assert.True(shapeText.Elements(ns + "tspan").Count() > 1);

            XElement connectorGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-connector-id") == connector.Id);
            XElement labelBackground = connectorGroup.Elements(ns + "rect")
                .Single(rect => (string?)rect.Attribute("data-officeimo-connector-label-background") == "true");
            Assert.True(double.Parse(labelBackground.Attribute("width")!.Value, CultureInfo.InvariantCulture) > 0D);
            Assert.Equal("0.902", labelBackground.Attribute("fill-opacity")!.Value);

            XElement labelText = connectorGroup.Elements(ns + "text").Single();
            Assert.True(labelText.Elements(ns + "tspan").Count() > 1);
        }

        [Fact]
        public void SvgRendererDrawsStyledShapeTextBackgrounds() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Text Background").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1.6, 0.7, "Escalation window");
            shape.TextStyle = new VisioTextStyle {
                Size = 11,
                TextWidth = 1.2,
                TextHeight = 0.42,
                BackgroundColor = OfficeColor.FromRgb(255, 236, 179),
                BackgroundTransparency = 25
            };

            string svg = page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 100 });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement background = shapeGroup.Elements(ns + "rect")
                .Single(rect => (string?)rect.Attribute("data-officeimo-text-background") == "true");
            Assert.Null(background.Attribute("data-officeimo-connector-label-background"));
            Assert.Equal("#FFECB3", background.Attribute("fill")!.Value);
            Assert.Equal("0.749", background.Attribute("fill-opacity")!.Value);
        }

        [Fact]
        public void SvgRendererPreservesStyledTextOpacity() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Text Opacity").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1.6, 0.7, "Escalation");
            shape.TextStyle = new VisioTextStyle {
                Size = 11,
                TextWidth = 1.2,
                TextHeight = 0.42,
                Color = OfficeColor.FromRgba(220, 38, 38, 128)
            };

            string svg = page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 100 });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement text = shapeGroup.Elements(ns + "text").Single();
            Assert.Equal("#DC2626", text.Attribute("fill")!.Value);
            Assert.Equal("0.502", text.Attribute("fill-opacity")!.Value);
        }

        [Fact]
        public void SvgRendererPreservesStyledTextUnderline() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Text Underline").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1.8, 0.8, "OfficeIMO");
            shape.TextStyle = new VisioTextStyle {
                Size = 18,
                TextWidth = 1.6,
                TextHeight = 0.5,
                Underline = true,
                Color = OfficeColor.FromRgb(22, 101, 52)
            };

            string svg = page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 100 });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement text = shapeGroup.Elements(ns + "text").Single();
            Assert.Equal("underline", text.Attribute("text-decoration")!.Value);
        }

        [Fact]
        public void SvgRendererPreservesStyledTextItalic() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Text Italic").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1.8, 0.8, "OfficeIMO");
            shape.TextStyle = new VisioTextStyle {
                Size = 18,
                TextWidth = 1.6,
                TextHeight = 0.5,
                Italic = true,
                Color = OfficeColor.FromRgb(22, 101, 52)
            };

            string svg = page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 100 });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement text = shapeGroup.Elements(ns + "text").Single();
            Assert.Equal("italic", text.Attribute("font-style")!.Value);
        }

        [Fact]
        public void SvgRendererRotatesStyledTextWithTextAngle() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Text Rotation").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1.8, 0.8, "OfficeIMO");
            shape.TextStyle = new VisioTextStyle {
                Size = 18,
                TextWidth = 1.6,
                TextHeight = 0.5,
                TextAngle = Math.PI / 4D,
                Color = OfficeColor.FromRgb(22, 101, 52)
            };

            string svg = page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 100 });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement text = shapeGroup.Elements(ns + "text").Single();
            Assert.StartsWith("rotate(-45", text.Attribute("transform")!.Value, StringComparison.Ordinal);
        }

        [Fact]
        public void SvgRendererRotatesStyledTextBackgroundWithTextAngle() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Text Background Rotation").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1.8, 0.8, "OfficeIMO");
            shape.TextStyle = new VisioTextStyle {
                Size = 18,
                TextWidth = 1.6,
                TextHeight = 0.5,
                TextAngle = Math.PI / 4D,
                BackgroundColor = OfficeColor.FromRgb(220, 38, 38),
                BackgroundTransparency = 0,
                Color = OfficeColor.FromRgb(22, 101, 52)
            };

            string svg = page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 100 });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement text = shapeGroup.Elements(ns + "text").Single();
            XElement background = shapeGroup.Elements(ns + "rect")
                .Single(rect => (string?)rect.Attribute("data-officeimo-text-background") == "true");
            Assert.Equal(text.Attribute("transform")!.Value, background.Attribute("transform")!.Value);
            Assert.Equal("#DC2626", background.Attribute("fill")!.Value);
        }

        [Fact]
        public void SvgRendererNudgesConnectorLabelsAwayFromShapeCollisions() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Label Avoidance").Size(6, 3);
            VisioShape source = page.AddRectangle(1, 1.5, 1, 0.6, "Source");
            VisioShape target = page.AddRectangle(5, 1.5, 1, 0.6, "Target");
            page.AddRectangle(3, 1.5, 1.1, 0.7, "Obstacle");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            connector.Label = "handoff";
            connector.PlaceLabel(0.5, width: 1.2, height: 0.3);

            string svg = page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 100 });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement connectorGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-connector-id") == connector.Id);
            XElement labelBackground = connectorGroup.Elements(ns + "rect")
                .Single(rect => (string?)rect.Attribute("data-officeimo-connector-label-background") == "true");
            XElement labelText = connectorGroup.Elements(ns + "text").Single();
            Assert.Equal("true", labelBackground.Attribute("data-officeimo-label-adjusted")!.Value);
            Assert.Equal("true", labelText.Attribute("data-officeimo-label-adjusted")!.Value);
            Assert.NotEqual("135", labelBackground.Attribute("y")!.Value);
        }

        [Fact]
        public void SvgRendererNudgesConnectorLabelsAwayFromEndpointShapeCollisions() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Endpoint Label Avoidance").Size(6, 3);
            VisioShape source = page.AddRectangle(1, 1.5, 1, 0.6, "Source");
            VisioShape target = page.AddRectangle(5, 1.5, 1, 0.6, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            connector.Label = "endpoint collision";
            connector.PlaceLabel(0, width: 1.2, height: 0.3);

            string svg = page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 100 });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement connectorGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-connector-id") == connector.Id);
            XElement labelBackground = connectorGroup.Elements(ns + "rect")
                .Single(rect => (string?)rect.Attribute("data-officeimo-connector-label-background") == "true");
            XElement labelText = connectorGroup.Elements(ns + "text").Single();
            Assert.Equal("true", labelBackground.Attribute("data-officeimo-label-adjusted")!.Value);
            Assert.Equal("true", labelText.Attribute("data-officeimo-label-adjusted")!.Value);
        }

        [Fact]
        public void SvgRendererNudgesConnectorLabelsAwayFromOtherConnectorLines() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Connector Label Crossing Avoidance").Size(6, 3);
            VisioShape source = page.AddRectangle(1, 1.5, 1, 0.4, "Source");
            VisioShape target = page.AddRectangle(5, 1.5, 1, 0.4, "Target");
            VisioConnector labeled = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            labeled.Label = "handoff";
            labeled.PlaceLabel(0.5, width: 1.2, height: 0.3);

            VisioShape top = page.AddRectangle(3, 2.7, 0.5, 0.4, "Top");
            VisioShape bottom = page.AddRectangle(3, 0.3, 0.5, 0.4, "Bottom");
            page.AddConnector(top, bottom, ConnectorKind.Straight, VisioSide.Bottom, VisioSide.Top);

            string svg = page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 100 });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement connectorGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-connector-id") == labeled.Id);
            XElement labelBackground = connectorGroup.Elements(ns + "rect")
                .Single(rect => (string?)rect.Attribute("data-officeimo-connector-label-background") == "true");
            XElement labelText = connectorGroup.Elements(ns + "text").Single();
            Assert.Equal("true", labelBackground.Attribute("data-officeimo-label-adjusted")!.Value);
            Assert.Equal("true", labelText.Attribute("data-officeimo-label-adjusted")!.Value);
            Assert.NotEqual("240", labelBackground.Attribute("x")!.Value);
        }

        [Fact]
        public void SvgRendererKeepsDenseConnectorLabelsSeparated() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Dense Label Clearance").Size(6, 3);
            VisioShape lowerSource = page.AddRectangle(1, 1.5, 0.6, 0.25, string.Empty);
            VisioShape lowerTarget = page.AddRectangle(5, 1.5, 0.6, 0.25, string.Empty);
            VisioConnector lower = page.AddConnector(lowerSource, lowerTarget, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            lower.Label = "phase one";
            lower.PlaceLabel(0.5, width: 1.2, height: 0.3);

            VisioShape upperSource = page.AddRectangle(1, 1.82, 0.6, 0.25, string.Empty);
            VisioShape upperTarget = page.AddRectangle(5, 1.82, 0.6, 0.25, string.Empty);
            VisioConnector upper = page.AddConnector(upperSource, upperTarget, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            upper.Label = "phase two";
            upper.PlaceLabel(0.5, width: 1.2, height: 0.3);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement lowerGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-connector-id") == lower.Id);
            XElement upperGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-connector-id") == upper.Id);
            XElement lowerBackground = lowerGroup.Elements(ns + "rect")
                .Single(rect => (string?)rect.Attribute("data-officeimo-connector-label-background") == "true");
            XElement upperBackground = upperGroup.Elements(ns + "rect")
                .Single(rect => (string?)rect.Attribute("data-officeimo-connector-label-background") == "true");

            Assert.Null(lowerBackground.Attribute("data-officeimo-label-adjusted"));
            Assert.Equal("true", upperBackground.Attribute("data-officeimo-label-adjusted")!.Value);
            Assert.NotEqual("103", upperBackground.Attribute("y")!.Value);
        }

        [Fact]
        public void SvgRendererProjectsBuiltInStencilMetadataAsVectorArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Stencil Artwork").Size(3, 2);
            VisioShape shape = page.AddStencilShape(VisioStencils.SecurityIdentity, "sec.firewall", "firewall", 1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillPattern = 0;
            shape.LinePattern = 0;

            string svg = page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 100 });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement artwork = shapeGroup.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-officeimo-stencil-artwork") == "true");
            Assert.Equal("security", artwork.Attribute("data-officeimo-stencil-key")!.Value);
            Assert.Contains(artwork.Descendants(ns + "path"), path => ((string?)path.Attribute("d"))?.IndexOf("Z", StringComparison.Ordinal) >= 0);
        }

        [Fact]
        public void SvgRendererRotatesStencilMetadataArtworkWithShape() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Rotated Stencil Artwork").Size(3, 2);
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillPattern = 0;
            shape.LinePattern = 0;
            shape.Angle = Math.PI / 4D;
            shape.SetUserCell("OfficeIMO.StencilId", "event.bus", "STR");

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement artwork = shapeGroup.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-officeimo-stencil-artwork") == "true");
            Assert.Equal("event", artwork.Attribute("data-officeimo-stencil-key")!.Value);
            Assert.StartsWith("rotate(-45", artwork.Attribute("transform")!.Value, StringComparison.Ordinal);
        }

        [Fact]
        public void SvgRendererDoesNotProjectSequenceFragmentRegionsAsCloudArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Sequence Fragment").Size(5, 3);
            VisioShape fragment = page.AddRectangle(2.5, 1.5, 4, 2, string.Empty);
            fragment.FillPattern = 0;
            fragment.LinePattern = 0;
            fragment.SetUserCell("OfficeIMO.Kind", "SequenceFragment", "STR");
            fragment.SetUserCell("OfficeIMO.StencilId", "seq.fragment", "STR");
            fragment.SetUserCell("OfficeIMO.StencilName", "Combined Fragment", "STR");
            fragment.SetUserCell("OfficeIMO.StencilAliases", "alt;combined-fragment;critical;fragment;loop;opt;region", "STR");
            fragment.SetUserCell("OfficeIMO.StencilTags", "Rectangle;seq;Sequence Diagram", "STR");

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null
            });

            Assert.DoesNotContain("data-officeimo-stencil-artwork=\"true\"", svg);
            Assert.DoesNotContain("data-officeimo-stencil-key=\"cloud\"", svg);
        }

        [Fact]
        public void SvgRendererProjectsPackageBackedPngPreviewArtwork() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package Preview").Size(3, 2);
            VisioShape shape = AddPackagePreviewShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement image = shapeGroup.Elements(ns + "image")
                .Single(element => (string?)element.Attribute("data-officeimo-package-preview-artwork") == "true");
            Assert.StartsWith("data:image/png;base64,", image.Attribute("href")!.Value, StringComparison.Ordinal);
            Assert.DoesNotContain("data-officeimo-stencil-artwork=\"true\"", svg);
        }

        [Fact]
        public void SvgRendererSniffsPackageBackedPreviewArtworkWhenMetadataIsGeneric() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package Preview Sniff").Size(3, 2);
            VisioShape shape = AddPackagePreviewShape(page, "application/octet-stream", ".bin", "../media/blob1.bin");

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement image = shapeGroup.Elements(ns + "image")
                .Single(element => (string?)element.Attribute("data-officeimo-package-preview-artwork") == "true");
            Assert.StartsWith("data:image/png;base64,", image.Attribute("href")!.Value, StringComparison.Ordinal);
        }

        [Fact]
        public void SvgRendererNormalizesPackagePreviewContentTypeParameters() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package Preview Content Type").Size(3, 2);
            VisioShape shape = AddPackagePreviewShape(page, "image/png; charset=binary", ".bin", "../media/blob1.bin");

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement image = shapeGroup.Elements(ns + "image")
                .Single(element => (string?)element.Attribute("data-officeimo-package-preview-artwork") == "true");
            Assert.StartsWith("data:image/png;base64,", image.Attribute("href")!.Value, StringComparison.Ordinal);
        }

        [Fact]
        public void SvgRendererSniffsPackageBackedSvgPreviewArtworkWithXmlPreamble() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Package SVG Preview").Size(3, 2);
            byte[] svgPreview = Encoding.UTF8.GetBytes(
                "\uFEFF<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<!-- OfficeIMO package preview -->" +
                "<!DOCTYPE svg>" +
                "<?officeimo preview?>" +
                "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"8\" height=\"8\"><rect width=\"8\" height=\"8\" fill=\"#2563eb\"/></svg>");
            VisioShape shape = AddPackagePreviewShape(page, "application/octet-stream", ".bin", "../media/blob1.bin", svgPreview);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement image = shapeGroup.Elements(ns + "image")
                .Single(element => (string?)element.Attribute("data-officeimo-package-preview-artwork") == "true");
            Assert.StartsWith("data:image/svg+xml;base64,", image.Attribute("href")!.Value, StringComparison.Ordinal);
            Assert.DoesNotContain("data-officeimo-stencil-artwork=\"true\"", svg);
        }

        [Fact]
        public void SvgRendererRotatesPackageBackedPngPreviewArtworkWithShape() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Rotated Package Preview").Size(3, 2);
            VisioShape shape = AddPackagePreviewShape(page);
            shape.Angle = Math.PI / 4D;

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement image = shapeGroup.Elements(ns + "image")
                .Single(element => (string?)element.Attribute("data-officeimo-package-preview-artwork") == "true");
            Assert.StartsWith("rotate(-45", image.Attribute("transform")!.Value, StringComparison.Ordinal);
        }

        [Fact]
        public void SvgRendererAppliesParentTransformsToGroupChildren() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Groups").Size(5, 5);
            VisioShape group = new("group", 3, 2, 2, 2, string.Empty) { Type = "Group", FillPattern = 0 };
            VisioShape child = new("child", 1, 1, 1, 1, "Child");
            group.Children.Add(child);
            page.Shapes.Add(group);

            string svg = page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 100 });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement childGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == child.Id);
            XElement path = childGroup.Element(ns + "path")!;
            Assert.Equal("M 250 350 L 350 350 L 350 250 L 250 250 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererUsesPreservedRelativeShapeGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Geometry").Size(3, 2);
            VisioShape shape = AddRelativeTriangleGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");
            Assert.Equal("M 100 150 L 200 150 L 150 50 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererPreservesGeometrySubpathBreaks() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Subpaths").Size(3, 2);
            VisioShape shape = AddSubpathBreakGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");

            Assert.Equal("M 110 140 L 145 140 L 127.5 105 Z M 155 95 L 190 95 L 172.5 60 Z", path.Attribute("d")!.Value);
            Assert.Equal("evenodd", path.Attribute("fill-rule")!.Value);
            Assert.Equal("evenodd", path.Attribute("clip-rule")!.Value);
        }

        [Fact]
        public void SvgRendererPreservesHolesInPreservedGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Geometry Hole").Size(3, 2);
            VisioShape shape = AddDonutGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");

            Assert.Equal("M 110 140 L 190 140 L 190 60 L 110 60 Z M 135 115 L 165 115 L 165 85 L 135 85 Z", path.Attribute("d")!.Value);
            Assert.Equal("evenodd", path.Attribute("fill-rule")!.Value);
            Assert.Equal("evenodd", path.Attribute("clip-rule")!.Value);
            Assert.NotEqual("none", path.Attribute("fill")!.Value);
        }

        [Fact]
        public void SvgRendererLeavesNoFillOpenGeometryUnclosed() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Open Geometry").Size(3, 2);
            VisioShape shape = AddOpenNoFillGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");

            Assert.Equal("M 110 130 L 190 130 L 190 70", path.Attribute("d")!.Value);
            Assert.Equal("none", path.Attribute("fill")!.Value);
            Assert.NotEqual("none", path.Attribute("stroke")!.Value);
        }

        [Fact]
        public void SvgRendererSkipsDeletedPreservedGeometryRows() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Deleted Rows").Size(3, 2);
            VisioShape shape = AddDeletedGeometryRowShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");

            Assert.Equal("M 100 150 L 200 150 L 150 50 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererUsesPreservedMasterShapeGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Master Geometry").Size(3, 2);
            VisioShape shape = AddMasterBackedTriangleGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");
            Assert.Equal("M 100 150 L 200 150 L 150 50 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererEvaluatesPreservedGeometryWidthHeightFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Formula Geometry").Size(3, 2);
            VisioShape shape = AddFormulaGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");
            Assert.Equal("M 125 125 L 175 125 L 150 75 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererEvaluatesPreservedGeometryLocPinFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("LocPin Formula Geometry").Size(3, 2);
            VisioShape shape = AddLocPinFormulaGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");
            Assert.Equal("M 96 136 L 204 136 L 150 64 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererEvaluatesPreservedGeometryShapeTransformFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Shape Transform Formula Geometry").Size(3, 2);
            VisioShape shape = AddShapeTransformFormulaGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");
            Assert.Equal("M 100 150 L 200 150 L 150 50 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererEvaluatesPreservedGeometryMinMaxFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Min Max Formula Geometry").Size(3, 2);
            VisioShape shape = AddMinMaxFormulaGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");
            Assert.Equal("M 110 120 L 180 120 L 150 60 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererEvaluatesPreservedGeometryScalarMathFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Scalar Math Formula Geometry").Size(3, 2);
            VisioShape shape = AddScalarMathFormulaGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");
            Assert.Equal("M 110 120 L 180 120 L 150 60 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererEvaluatesPreservedGeometryTrigonometricFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Trig Formula Geometry").Size(3, 2);
            VisioShape shape = AddTrigonometricFormulaGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");
            Assert.Equal("M 110 120 L 180 120 L 150 60 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererEvaluatesPreservedGeometryAdvancedMathFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Advanced Math Formula Geometry").Size(3, 2);
            VisioShape shape = AddAdvancedMathFormulaGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");
            Assert.Equal("M 110 120 L 180 120 L 150 60 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererEvaluatesPreservedGeometryPowerOperatorFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Power Operator Formula Geometry").Size(3, 2);
            VisioShape shape = AddPowerOperatorFormulaGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");
            Assert.Equal("M 110 120 L 180 120 L 150 60 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererEvaluatesPreservedGeometryUnitFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Unit Formula Geometry").Size(3, 2);
            VisioShape shape = AddUnitFormulaGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");
            Assert.Equal("M 110 120 L 180 120 L 150 60 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererEvaluatesPreservedGeometryGuardedFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Guarded Formula Geometry").Size(3, 2);
            VisioShape shape = AddGuardedFormulaGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");
            Assert.Equal("M 125 125 L 175 125 L 150 75 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererEvaluatesPreservedGeometryIfFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("If Formula Geometry").Size(3, 2);
            VisioShape shape = AddIfFormulaGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");
            Assert.Equal("M 110 120 L 180 120 L 150 60 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererEvaluatesOnlySelectedPreservedGeometryIfBranches() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Lazy If Formula Geometry").Size(3, 2);
            VisioShape shape = AddLazyIfFormulaGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");
            Assert.Equal("M 110 120 L 180 120 L 150 60 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererEvaluatesPreservedGeometryLogicalIfFormulas() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Logical If Formula Geometry").Size(3, 2);
            VisioShape shape = AddLogicalIfFormulaGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");
            Assert.Equal("M 110 120 L 180 120 L 150 60 Z", path.Attribute("d")!.Value);
        }

        [Fact]
        public void SvgRendererRespectsPreservedGeometryVisibilityFlags() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Geometry Flags").Size(3, 2);
            VisioShape shape = AddGeometryFlagShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            List<XElement> paths = shapeGroup.Elements(ns + "path")
                .Where(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true")
                .ToList();

            Assert.Equal(2, paths.Count);
            Assert.Equal("none", paths[0].Attribute("fill")!.Value);
            Assert.NotEqual("none", paths[0].Attribute("stroke")!.Value);
            Assert.NotEqual("none", paths[1].Attribute("fill")!.Value);
            Assert.Equal("none", paths[1].Attribute("stroke")!.Value);
            Assert.DoesNotContain(paths, path => path.Attribute("d")!.Value == "M 100 150 L 200 150 L 200 50 L 100 50 Z");
        }

        [Fact]
        public void SvgRendererFlattensPreservedEllipseGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Ellipse Geometry").Size(3, 2);
            VisioShape shape = AddEllipseGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            string pathData = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true")
                .Attribute("d")!.Value;
            Assert.StartsWith("M 200 100 L ", pathData, StringComparison.Ordinal);
            Assert.Contains("150 50", pathData, StringComparison.Ordinal);
            Assert.Contains("100 100", pathData, StringComparison.Ordinal);
            Assert.True(pathData.Split(new[] { " L " }, StringSplitOptions.None).Length > 20);
        }

        [Fact]
        public void SvgRendererDrawsPreservedInfiniteLineGeometryAsOpenPath() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Infinite Line Geometry").Size(3, 2);
            VisioShape shape = AddInfiniteLineGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            XElement path = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");
            Assert.Equal("M 100 150 L 200 50", path.Attribute("d")!.Value);
            Assert.Equal("none", path.Attribute("fill")!.Value);
            Assert.NotEqual("none", path.Attribute("stroke")!.Value);
        }

        [Fact]
        public void SvgRendererExpandsPreservedPolylineToGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Polyline Geometry").Size(3, 2);
            VisioShape shape = AddPolylineGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            string pathData = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true")
                .Attribute("d")!.Value;
            Assert.Equal("M 150 50 L 200 100 L 150 150 L 100 100 Z", pathData);
        }

        [Fact]
        public void SvgRendererEvaluatesMinMaxInsidePreservedPolylineFormula() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Polyline Formula Geometry").Size(3, 2);
            VisioShape shape = AddPolylineMinMaxFormulaGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            string pathData = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true")
                .Attribute("d")!.Value;
            Assert.Equal("M 150 60 L 180 100 L 150 140 L 110 100 Z", pathData);
        }

        [Fact]
        public void SvgRendererEvaluatesPercentageLiteralsInsidePreservedPolylineFormula() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Polyline Percent Formula Geometry").Size(3, 2);
            VisioShape shape = AddPolylinePercentageFormulaGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            string pathData = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true")
                .Attribute("d")!.Value;
            Assert.Equal("M 150 60 L 180 100 L 150 140 L 120 100 Z", pathData);
        }

        [Fact]
        public void SvgRendererFlattensPreservedArcToGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Arc Geometry").Size(3, 2);
            VisioShape shape = AddArcGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            string pathData = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true")
                .Attribute("d")!.Value;
            Assert.StartsWith("M 100 150 L ", pathData, StringComparison.Ordinal);
            Assert.Contains(" L 200 150", pathData, StringComparison.Ordinal);
            Assert.True(pathData.Split(new[] { " L " }, StringSplitOptions.None).Length > 10);
        }

        [Fact]
        public void SvgRendererFlattensPreservedEllipticalArcToGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Elliptical Arc Geometry").Size(3, 2);
            VisioShape shape = AddEllipticalArcGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            string pathData = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true")
                .Attribute("d")!.Value;
            Assert.StartsWith("M 100 150 L ", pathData, StringComparison.Ordinal);
            Assert.Contains(" L 200 150", pathData, StringComparison.Ordinal);
            Assert.Contains("150 50", pathData, StringComparison.Ordinal);
            Assert.True(pathData.Split(new[] { " L " }, StringSplitOptions.None).Length > 10);
        }

        [Fact]
        public void SvgRendererFlattensPreservedRelativeEllipticalArcToGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Relative Elliptical Arc Geometry").Size(3, 2);
            VisioShape shape = AddRelativeEllipticalArcGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            string pathData = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true")
                .Attribute("d")!.Value;
            Assert.StartsWith("M 100 150 L ", pathData, StringComparison.Ordinal);
            Assert.Contains(" L 200 150", pathData, StringComparison.Ordinal);
            Assert.Contains("150 50", pathData, StringComparison.Ordinal);
            Assert.True(pathData.Split(new[] { " L " }, StringSplitOptions.None).Length > 10);
        }

        [Fact]
        public void SvgRendererFlattensPreservedRelativeCubicBezierGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Cubic Bezier Geometry").Size(3, 2);
            VisioShape shape = AddRelativeCubicBezierGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            string pathData = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true")
                .Attribute("d")!.Value;
            Assert.StartsWith("M 100 150 L ", pathData, StringComparison.Ordinal);
            Assert.Contains(" L 200 150", pathData, StringComparison.Ordinal);
            Assert.Contains("150 75", pathData, StringComparison.Ordinal);
            Assert.True(pathData.Split(new[] { " L " }, StringSplitOptions.None).Length > 10);
        }

        [Fact]
        public void SvgRendererFlattensPreservedAbsoluteCubicBezierGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Absolute Cubic Bezier Geometry").Size(3, 2);
            VisioShape shape = AddAbsoluteCubicBezierGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            string pathData = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true")
                .Attribute("d")!.Value;
            Assert.StartsWith("M 100 150 L ", pathData, StringComparison.Ordinal);
            Assert.Contains(" L 200 150", pathData, StringComparison.Ordinal);
            Assert.Contains("150 75", pathData, StringComparison.Ordinal);
            Assert.True(pathData.Split(new[] { " L " }, StringSplitOptions.None).Length > 10);
        }

        [Fact]
        public void SvgRendererFlattensPreservedRelativeQuadraticBezierGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Quadratic Bezier Geometry").Size(3, 2);
            VisioShape shape = AddRelativeQuadraticBezierGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            string pathData = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true")
                .Attribute("d")!.Value;
            Assert.StartsWith("M 100 150 L ", pathData, StringComparison.Ordinal);
            Assert.Contains(" L 200 150", pathData, StringComparison.Ordinal);
            Assert.Contains("150 100", pathData, StringComparison.Ordinal);
            Assert.True(pathData.Split(new[] { " L " }, StringSplitOptions.None).Length > 10);
        }

        [Fact]
        public void SvgRendererFlattensPreservedAbsoluteQuadraticBezierGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Absolute Quadratic Bezier Geometry").Size(3, 2);
            VisioShape shape = AddAbsoluteQuadraticBezierGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            string pathData = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true")
                .Attribute("d")!.Value;
            Assert.StartsWith("M 100 150 L ", pathData, StringComparison.Ordinal);
            Assert.Contains(" L 200 150", pathData, StringComparison.Ordinal);
            Assert.Contains("150 100", pathData, StringComparison.Ordinal);
            Assert.True(pathData.Split(new[] { " L " }, StringSplitOptions.None).Length > 10);
        }

        [Fact]
        public void SvgRendererFlattensPreservedSplineGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Spline Geometry").Size(3, 2);
            VisioShape shape = AddSplineGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            string pathData = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true")
                .Attribute("d")!.Value;
            Assert.StartsWith("M 100 150 L ", pathData, StringComparison.Ordinal);
            Assert.Contains("125 50", pathData, StringComparison.Ordinal);
            Assert.Contains("175 50", pathData, StringComparison.Ordinal);
            Assert.Contains(" L 200 150", pathData, StringComparison.Ordinal);
            Assert.True(pathData.Split(new[] { " L " }, StringSplitOptions.None).Length > 30);
        }

        [Fact]
        public void SvgRendererSkipsDeletedPreservedSplineKnotRows() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Deleted Spline Knot Geometry").Size(3, 2);
            VisioShape shape = AddDeletedSplineKnotGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            string pathData = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true")
                .Attribute("d")!.Value;
            Assert.Equal("M 100 150 L 200 150 L 200 50 Z", pathData);
            Assert.DoesNotContain("100 50", pathData, StringComparison.Ordinal);
        }

        [Fact]
        public void SvgRendererFlattensPreservedNurbsGeometry() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved NURBS Geometry").Size(3, 2);
            VisioShape shape = AddNurbsGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            string pathData = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true")
                .Attribute("d")!.Value;
            Assert.StartsWith("M 100 150 L ", pathData, StringComparison.Ordinal);
            Assert.Contains("150 75", pathData, StringComparison.Ordinal);
            Assert.Contains(" L 200 150", pathData, StringComparison.Ordinal);
            Assert.True(pathData.Split(new[] { " L " }, StringSplitOptions.None).Length > 10);
        }

        [Fact]
        public void SvgRendererUsesVisioCompactNurbsKnotVector() {
            using MemoryStream packageStream = new();
            VisioDocument document = VisioDocument.Create(packageStream);
            VisioPage page = document.AddPage("Preserved Non Uniform NURBS Geometry").Size(3, 2);
            VisioShape shape = AddNonUniformNurbsGeometryShape(page);

            string svg = page.ToSvg(new VisioSvgSaveOptions {
                BackgroundColor = null,
                PixelsPerInch = 100
            });

            XDocument parsed = XDocument.Parse(svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement shapeGroup = parsed.Root!.Descendants(ns + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == shape.Id);
            string pathData = shapeGroup.Elements(ns + "path")
                .Single(element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true")
                .Attribute("d")!.Value;
            Assert.StartsWith("M 100 150 L ", pathData, StringComparison.Ordinal);
            Assert.Contains("128.75 50", pathData, StringComparison.Ordinal);
            Assert.Contains(" L 200 150", pathData, StringComparison.Ordinal);
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

        private static VisioShape AddPackagePreviewShape(
            VisioPage page,
            string contentType = "image/png",
            string extension = ".png",
            string target = "../media/image1.png",
            byte[]? data = null) {
            VisioMaster master = new("package-master", "FancyCloud", new VisioShape("master-shape", 0, 0, 1, 1, string.Empty));
            master.RawMasterRelationships.Add(new VisioAssets.MasterRelationshipContent {
                Id = "rIdImage",
                Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                Target = target,
                ContentType = contentType,
                Extension = extension,
                Data = data ?? Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgSPj/HwAEIgJfhz+lZwAAAABJRU5ErkJggg==")
            });
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillPattern = 0;
            shape.LinePattern = 0;
            shape.Master = master;
            shape.NameU = master.NameU;
            shape.SetUserCell("OfficeIMO.StencilId", "package.fancy-cloud", "STR");
            shape.SetUserCell("OfficeIMO.StencilPreviewImageRelationshipId", "rIdImage", "STR");
            shape.SetUserCell("OfficeIMO.StencilPreviewImageTarget", target, "STR");
            shape.SetUserCell("OfficeIMO.StencilPreviewImageContentType", contentType, "STR");
            shape.SetUserCell("OfficeIMO.StencilPreviewImageExtension", extension, "STR");
            return shape;
        }

        private static VisioShape AddRelativeTriangleGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateRelativeTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddSubpathBreakGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateSubpathBreakGeometrySection());
            return shape;
        }

        private static VisioShape AddDonutGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateDonutGeometrySection());
            return shape;
        }

        private static VisioShape AddOpenNoFillGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineWeight = 0.04D;
            shape.PreservedGeometrySections.Add(CreateOpenNoFillGeometrySection());
            return shape;
        }

        private static VisioShape AddDeletedGeometryRowShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateDeletedRowGeometrySection());
            return shape;
        }

        private static VisioShape AddArcGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateArcGeometrySection());
            return shape;
        }

        private static VisioShape AddEllipticalArcGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateEllipticalArcGeometrySection());
            return shape;
        }

        private static VisioShape AddRelativeEllipticalArcGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateRelativeEllipticalArcGeometrySection());
            return shape;
        }

        private static VisioShape AddRelativeCubicBezierGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateRelativeCubicBezierGeometrySection());
            return shape;
        }

        private static VisioShape AddAbsoluteCubicBezierGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateAbsoluteCubicBezierGeometrySection());
            return shape;
        }

        private static VisioShape AddRelativeQuadraticBezierGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateRelativeQuadraticBezierGeometrySection());
            return shape;
        }

        private static VisioShape AddAbsoluteQuadraticBezierGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateAbsoluteQuadraticBezierGeometrySection());
            return shape;
        }

        private static VisioShape AddSplineGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateSplineGeometrySection());
            return shape;
        }

        private static VisioShape AddDeletedSplineKnotGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateDeletedSplineKnotGeometrySection());
            return shape;
        }

        private static VisioShape AddNurbsGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateNurbsGeometrySection());
            return shape;
        }

        private static VisioShape AddNonUniformNurbsGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateNonUniformNurbsGeometrySection());
            return shape;
        }

        private static VisioShape AddMasterBackedTriangleGeometryShape(VisioPage page) {
            VisioShape masterShape = new("master-shape", 1, 0.5, 2, 1, string.Empty);
            masterShape.PreservedGeometrySections.Add(CreateRelativeTriangleGeometrySection());
            VisioMaster master = new("master-triangle", "PackageTriangle", masterShape);
            VisioShape shape = page.AddShape("master-backed-triangle", master, 1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            return shape;
        }

        private static VisioShape AddFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddLocPinFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateLocPinFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddShapeTransformFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateShapeTransformFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddMinMaxFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateMinMaxFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddScalarMathFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateScalarMathFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddTrigonometricFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateTrigonometricFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddAdvancedMathFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateAdvancedMathFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddPowerOperatorFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreatePowerOperatorFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddUnitFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateUnitFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddGuardedFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateGuardedFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddIfFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateIfFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddLazyIfFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateLazyIfFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddLogicalIfFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateLogicalIfFormulaTriangleGeometrySection());
            return shape;
        }

        private static VisioShape AddGeometryFlagShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.LineWeight = 0.03D;
            shape.PreservedGeometrySections.Add(CreateRelativeTriangleGeometrySection(noFill: true));
            shape.PreservedGeometrySections.Add(CreateInsetTriangleGeometrySection(noLine: true));
            shape.PreservedGeometrySections.Add(CreateFullRectangleGeometrySection(noShow: true));
            return shape;
        }

        private static VisioShape AddEllipseGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreateEllipseGeometrySection());
            return shape;
        }

        private static VisioShape AddInfiniteLineGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(37, 99, 235);
            shape.LineColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineWeight = 0.05D;
            shape.PreservedGeometrySections.Add(CreateInfiniteLineGeometrySection());
            return shape;
        }

        private static VisioShape AddPolylineGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1, 1, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreatePolylineGeometrySection());
            return shape;
        }

        private static VisioShape AddPolylineMinMaxFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreatePolylineMinMaxFormulaGeometrySection());
            return shape;
        }

        private static VisioShape AddPolylinePercentageFormulaGeometryShape(VisioPage page) {
            VisioShape shape = page.AddRectangle(1.5, 1, 1.2, 0.8, string.Empty);
            shape.FillColor = OfficeColor.FromRgb(220, 38, 38);
            shape.LineColor = OfficeColor.FromRgb(127, 29, 29);
            shape.PreservedGeometrySections.Add(CreatePolylinePercentageFormulaGeometrySection());
            return shape;
        }

        private static XElement CreateRelativeTriangleGeometrySection() {
            return CreateRelativeTriangleGeometrySection(noFill: false);
        }

        private static XElement CreateRelativeTriangleGeometrySection(bool noFill = false, bool noLine = false, bool noShow = false) {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                CreateGeometryRow(ns, noFill, noLine, noShow),
                new XElement(ns + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateSubpathBreakGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                CreateGeometryRow(ns, noFill: false, noLine: false, noShow: false),
                new XElement(ns + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.45")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.275")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.45"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "5"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.55")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.55"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "6"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.9")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.55"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "7"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.725")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.9"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "8"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.55")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.55"))));
        }

        private static XElement CreateDonutGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                CreateGeometryRow(ns, noFill: false, noLine: false, noShow: false),
                new XElement(ns + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.9")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.9")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.9"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.9"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "5"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "6"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.35")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.35"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "7"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.65")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.35"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "8"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.65")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.65"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "9"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.35")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.65"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "10"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.35")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.35"))));
        }

        private static XElement CreateOpenNoFillGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                CreateGeometryRow(ns, noFill: true, noLine: false, noShow: false),
                new XElement(ns + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.2"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.9")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.2"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.9")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.8"))));
        }

        private static XElement CreateDeletedRowGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                CreateGeometryRow(ns, noFill: false, noLine: false, noShow: false),
                new XElement(ns + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "3"), new XAttribute("Del", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "5"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateInsetTriangleGeometrySection(bool noFill = false, bool noLine = false, bool noShow = false) {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "1"),
                CreateGeometryRow(ns, noFill, noLine, noShow),
                new XElement(ns + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.75")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.75"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.25"))));
        }

        private static XElement CreateFullRectangleGeometrySection(bool noFill = false, bool noLine = false, bool noShow = false) {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "2"),
                CreateGeometryRow(ns, noFill, noLine, noShow),
                new XElement(ns + "Row", new XAttribute("T", "RelMoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "5"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateGeometryRow(XNamespace ns, bool noFill, bool noLine, bool noShow) {
            return new XElement(ns + "Row", new XAttribute("T", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Cell", new XAttribute("N", "NoFill"), new XAttribute("V", noFill ? "1" : "0")),
                new XElement(ns + "Cell", new XAttribute("N", "NoLine"), new XAttribute("V", noLine ? "1" : "0")),
                new XElement(ns + "Cell", new XAttribute("N", "NoShow"), new XAttribute("V", noShow ? "1" : "0")));
        }

        private static XElement CreateFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "=Width * 0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "=Height * 0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "Width * 0.75")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "Height / 4"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Width / 2")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "Height * 0.75"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "=Width * 0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "Height * (1 / 4)"))));
        }

        private static XElement CreateLocPinFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "LocPinX - Width * 0.45")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "LocPinY - Height * 0.45"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "LocPinX + Width * 0.45")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "LocPinY - Height * 0.45"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "LocPinX")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "LocPinY + Height * 0.45"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "LocPinX - Width * 0.45")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "LocPinY - Height * 0.45"))));
        }

        private static XElement CreateShapeTransformFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "PinX - PinX")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "PinY - PinY"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "PinX - 0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "PinY - PinY"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "PinX - Width")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "IF(Angle=0, PinY, 0)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "PinX - PinX")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "PinY - PinY"))));
        }

        private static XElement CreateMinMaxFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "=MIN(Width, Height) * 0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "=Height * 0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "MAX(Width, Height) * 0.75")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "MIN(Width, Height) * 0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Width / 2")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "MIN(Width, Height)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "MIN(Width, Height) * 0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "Height * 0.25"))));
        }

        private static XElement CreateScalarMathFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "=ABS(-Height * 0.25)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "SQRT(Height * Height) * 0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "SQRT(Width * Width) * 0.75")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "ABS(-Height / 4)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Width / 2")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "SQRT(ABS(-Height) * ABS(-Height))"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "ABS(Height * -0.25)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "SQRT(Height * Height) / 4"))));
        }

        private static XElement CreateTrigonometricFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "(SIN(PI() - PI()) + COS(0)) * Height * 0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "COS(PI() - PI()) * Height * 0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "COS(0) * Width * 0.75")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "(SIN(0) + COS(0)) * Height * 0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Width / 2")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "COS(PI() - PI()) * Height"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "(SIN(0) + COS(0)) * Height * 0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "COS(0) * Height * 0.25"))));
        }

        private static XElement CreateAdvancedMathFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "(POW(2, 0) + TAN(ATAN2(0, 1)) + TAN(ATAN(0))) * Height * 0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "(RAD(DEG(0)) + 1) * Height * 0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "ROUND(Width * 0.749, 2)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "INT(Height * 0.9) + Height * 0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "POW(Width, 1) / 2")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "ROUND(Height * 0.99, 1)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "POW(Height, 1) * 0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "(INT(1.9) + TAN(ATAN2(0, 1))) * Height * 0.25"))));
        }

        private static XElement CreateUnitFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "TAN(45 deg) * 0.2 in")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "COS(0 rad) * 0.2 in"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "22.86 mm")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "2.54 cm / 5"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Width / (1 in + 1)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "Height * (1 ft / 12)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.2 in")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "2.54 cm / 5"))));
        }

        private static XElement CreatePowerOperatorFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "(Height ^ 2) / Height * 25%")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "(Height ^ 2) / (Height * 4)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "Width * 75%")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "Height ^ 0 * Height / 4"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Width / 2")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "Height ^ 1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Height ^ 2 / Height * 25%")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "Height ^ 0 * Height / 4"))));
        }

        private static XElement CreateGuardedFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "=GUARD(Width * 0.25)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "GUARD(Height * 0.25)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "GUARD(Width * 0.75)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "GUARD(Height / 4)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "GUARD(Width / 2)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "GUARD(Height * 0.75)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "GUARD(Width * 0.25)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "GUARD(Height * (1 / 4))"))));
        }

        private static XElement CreateIfFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(Width > Height, Height * 0.25, Width * 0.25)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "IF(Width < Height, Width * 0.25, Height * 0.25)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "IF(Width >= Height, Width * 0.75, Height * 0.75)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "IF(FALSE, Width, Height * 0.25)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(Width = Height, Width, Width / 2)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "IF(TRUE, Height, Width)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(Width <> Height, Height * 0.25, Width * 0.25)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "IF(Width != Height, Height * 0.25, Width * 0.25)"))));
        }

        private static XElement CreateLazyIfFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(FALSE, Width / 0, Height * 0.25)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "IF(TRUE, Height * 0.25, Width / 0)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "IF(TRUE, Width * 0.75, Height / 0)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "IF(FALSE, Width / 0, Height * 0.25)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(Height, Width / 2, Width / 0)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "IF(Width > Height, IF(TRUE, Height, Width / 0), Width / 0)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(FALSE, Width / 0, Height * 0.25)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "IF(TRUE, MIN(Height, Width) * 0.25, Width / 0)"))));
        }

        private static XElement CreateLogicalIfFormulaTriangleGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(AND(Width > Height, NOT(FALSE)), Height * 0.25, Width / 0)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "IF(OR(FALSE, Height > 0), Height * 0.25, Width / 0)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("F", "IF(AND(TRUE, Width >= Height, NOT(Height > Width)), Width * 0.75, Width / 0)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "IF(OR(Width < Height, NOT(FALSE)), Height * 0.25, Width / 0)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(NOT(Width = Height), Width / 2, Width / 0)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "IF(AND(OR(FALSE, TRUE), Width > Height), Height, Width / 0)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "IF(OR(FALSE, Width <> Height), Height * 0.25, Width / 0)")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("F", "IF(AND(TRUE, NOT(Width < Height)), Height * 0.25, Width / 0)"))));
        }

        private static XElement CreateEllipseGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "Ellipse"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "C"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "1"))));
        }

        private static XElement CreateInfiniteLineGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "InfiniteLine"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1"))));
        }

        private static XElement CreatePolylineGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "PolylineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("F", "POLYLINE(0,0,0.5,1,1,0.5,0.5,0,0,0.5,0.5,1)"))));
        }

        private static XElement CreatePolylineMinMaxFormulaGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Width / 2")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "MIN(Width, Height)"))),
                new XElement(ns + "Row", new XAttribute("T", "PolylineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "Width / 2")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "MIN(Width, Height)")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("F", "POLYLINE(1,1,Width/2,MIN(Width,Height),MAX(Width,Height)*0.75,MIN(Width,Height)/2,Width/2,0,MIN(Width,Height)*0.25,MIN(Width,Height)/2,Width/2,MIN(Width,Height))"))));
        }

        private static XElement CreatePolylinePercentageFormulaGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "50% * Width")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "100% * Height"))),
                new XElement(ns + "Row", new XAttribute("T", "PolylineTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "50% * Width")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "100% * Height")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("F", "POLYLINE(1,1,50%*Width,100%*Height,75%*Width,50%*Height,50%*Width,0%*Height,25%*Width,50%*Height,50%*Width,100%*Height)"))));
        }

        private static XElement CreateArcGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "ArcTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.25"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "5"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateEllipticalArcGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "EllipticalArcTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "C"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateRelativeEllipticalArcGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelEllipticalArcTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "C"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateRelativeCubicBezierGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelCubBezTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "C"), new XAttribute("V", "0.75")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateAbsoluteCubicBezierGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "CubBezTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "C"), new XAttribute("V", "0.75")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateRelativeQuadraticBezierGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "RelQuadBezTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateAbsoluteQuadraticBezierGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "QuadBezTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.5")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateSplineGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "SplineStart"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.25")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "C"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "3"))),
                new XElement(ns + "Row", new XAttribute("T", "SplineKnot"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.75")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.5"))),
                new XElement(ns + "Row", new XAttribute("T", "SplineKnot"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "5"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateDeletedSplineKnotGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "SplineStart"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "SplineKnot"), new XAttribute("IX", "3"), new XAttribute("Del", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.5"))),
                new XElement(ns + "Row", new XAttribute("T", "SplineKnot"), new XAttribute("IX", "4"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "1"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "5"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateNurbsGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "NURBSTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "C"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "E"), new XAttribute("F", "NURBS(1,3,0,0,0.25,1,0,1,0.75,1,0,1)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }

        private static XElement CreateNonUniformNurbsGeometrySection() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return new XElement(ns + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                new XElement(ns + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))),
                new XElement(ns + "Row", new XAttribute("T", "NURBSTo"), new XAttribute("IX", "2"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "B"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "C"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "D"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "E"), new XAttribute("F", "NURBS(2,2,0,0,0.25,1,0,1,0.5,1,0,1,0.75,0,0.15,1)"))),
                new XElement(ns + "Row", new XAttribute("T", "LineTo"), new XAttribute("IX", "3"),
                    new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0"))));
        }
    }
}
