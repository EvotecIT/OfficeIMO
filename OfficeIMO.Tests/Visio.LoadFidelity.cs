using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioLoadFidelity {
        [Fact]
        public void LoadDetectsStraightConnectorWhenGeometryHasHeaderRow() {
            string filePath = CreateConnectorDocument(ConnectorKind.Straight);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement connectorShape = GetConnectorShape(pageDoc, ns);
                XElement geometry = connectorShape.Elements(ns + "Section").First(section => (string?)section.Attribute("N") == "Geometry");
                geometry.AddFirst(
                    new XElement(ns + "Row",
                        new XAttribute("T", "Geometry"),
                        new XElement(ns + "Cell", new XAttribute("N", "NoFill"), new XAttribute("V", "0")),
                        new XElement(ns + "Cell", new XAttribute("N", "NoLine"), new XAttribute("V", "0"))));
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector connector = Assert.Single(loaded.Pages[0].Connectors);
            Assert.Equal(ConnectorKind.Straight, connector.Kind);
        }

        [Fact]
        public void LoadDetectsRightAngleConnectorFromOrthogonalPolyline() {
            string filePath = CreateConnectorDocument(ConnectorKind.RightAngle);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement connectorShape = GetConnectorShape(pageDoc, ns);
                XElement geometry = connectorShape.Elements(ns + "Section").First(section => (string?)section.Attribute("N") == "Geometry");
                geometry.RemoveNodes();
                geometry.Add(
                    new XElement(ns + "Row",
                        new XAttribute("T", "Geometry"),
                        new XElement(ns + "Cell", new XAttribute("N", "NoFill"), new XAttribute("V", "0"))),
                    CreateRow(ns, "MoveTo", 1, 1),
                    CreateRow(ns, "LineTo", 1, 3),
                    CreateRow(ns, "LineTo", 4, 3),
                    CreateRow(ns, "LineTo", 4, 2));
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector connector = Assert.Single(loaded.Pages[0].Connectors);
            Assert.Equal(ConnectorKind.RightAngle, connector.Kind);
        }

        [Fact]
        public void LoadDetectsCurvedConnectorFromArcGeometry() {
            string filePath = CreateConnectorDocument(ConnectorKind.Straight);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement connectorShape = GetConnectorShape(pageDoc, ns);
                XElement geometry = connectorShape.Elements(ns + "Section").First(section => (string?)section.Attribute("N") == "Geometry");
                geometry.RemoveNodes();
                geometry.Add(
                    CreateRow(ns, "MoveTo", 1, 1),
                    new XElement(ns + "Row",
                        new XAttribute("T", "ArcTo"),
                        new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "4")),
                        new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "2")),
                        new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.5"))));
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector connector = Assert.Single(loaded.Pages[0].Connectors);
            Assert.Equal(ConnectorKind.Curved, connector.Kind);
        }

        [Fact]
        public void ThemeXmlRoundTripsWithoutLosingCustomContent() {
            string filePath = CreateThemedDocument();
            string originalThemeXml = """
                <?xml version="1.0" encoding="utf-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
                  <a:themeElements>
                    <a:clrScheme name="Custom Colors" />
                  </a:themeElements>
                  <a:objectDefaults />
                </a:theme>
                """;
            RewriteEntry(filePath, "visio/theme/theme1.xml", originalThemeXml);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.True(XNode.DeepEquals(
                NormalizeTheme(originalThemeXml),
                NormalizeTheme(ReadEntry(savedPath, "visio/theme/theme1.xml"))));
        }

        [Fact]
        public void ThemeNameCanChangeWithoutDroppingCustomThemeStructure() {
            string filePath = CreateThemedDocument();
            RewriteEntry(filePath, "visio/theme/theme1.xml", """
                <?xml version="1.0" encoding="utf-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
                  <a:themeElements>
                    <a:fmtScheme name="Formatting" />
                  </a:themeElements>
                </a:theme>
                """);

            VisioDocument loaded = VisioDocument.Load(filePath);
            loaded.Theme!.Name = "Renamed Theme";

            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            XDocument savedTheme = XDocument.Parse(ReadEntry(savedPath, "visio/theme/theme1.xml"));
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            Assert.Equal("Renamed Theme", savedTheme.Root?.Attribute("name")?.Value);
            Assert.NotNull(savedTheme.Root?.Element(a + "themeElements"));
            Assert.NotNull(savedTheme.Root?.Element(a + "themeElements")?.Element(a + "fmtScheme"));
        }

        [Fact]
        public void ShapeChildOrderIsPreservedOnRoundTrip() {
            string filePath = CreateShapeDocument();

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement shape = GetFirstShape(pageDoc, ns);
                XElement pinX = shape.Elements(ns + "Cell")
                    .First(cell => (string?)cell.Attribute("N") == "PinX");
                XElement lineWeight = shape.Elements(ns + "Cell")
                    .First(cell => (string?)cell.Attribute("N") == "LineWeight");
                XElement geometry = shape.Elements(ns + "Section")
                    .First(section => (string?)section.Attribute("N") == "Geometry");
                XElement sidecarElement = new(ns + "ShapeMeta",
                    new XAttribute("KeepOrder", "1"));

                XElement[] remainingChildren = shape.Elements()
                    .Where(element => element != pinX && element != lineWeight && element != geometry)
                    .ToArray();

                shape.RemoveNodes();
                shape.Add(lineWeight, sidecarElement, geometry, pinX);
                foreach (XElement remainingChild in remainingChildren) {
                    shape.Add(remainingChild);
                }
            });

            string originalOrder = ReadFirstShapeChildOrder(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalOrder, ReadFirstShapeChildOrder(savedPath));
        }

        [Fact]
        public void DocumentRootAttributesAndElementsArePreservedOnRoundTrip() {
            string filePath = CreateShapeDocument();

            RewriteEntry(filePath, "visio/document.xml", """
                <?xml version="1.0" encoding="utf-8"?>
                <VisioDocument xmlns="http://schemas.microsoft.com/office/visio/2012/main" CustomDocFlag="KeepMe">
                  <PreviewScope Mode="Expanded" />
                  <DocumentSettings TopPage="0" DefaultTextStyle="0" DefaultLineStyle="0" DefaultFillStyle="0" DefaultGuideStyle="4">
                    <GlueSettings>9</GlueSettings>
                    <SnapSettings>295</SnapSettings>
                    <SnapExtensions>34</SnapExtensions>
                    <SnapAngles />
                    <DynamicGridEnabled>1</DynamicGridEnabled>
                    <ProtectStyles>0</ProtectStyles>
                    <ProtectShapes>0</ProtectShapes>
                    <ProtectMasters>0</ProtectMasters>
                    <ProtectBkgnds>0</ProtectBkgnds>
                  </DocumentSettings>
                  <Colors />
                  <FaceNames />
                  <StyleSheets>
                    <StyleSheet ID="0" Name="No Style" NameU="No Style">
                      <Cell N="EnableLineProps" V="1" />
                    </StyleSheet>
                  </StyleSheets>
                </VisioDocument>
                """);

            string originalFragments = ReadDocumentRootFragments(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalFragments, ReadDocumentRootFragments(savedPath));
        }

        [Fact]
        public void DocumentSettingsAttributesAndElementsArePreservedOnRoundTrip() {
            string filePath = CreateShapeDocument();

            RewriteEntry(filePath, "visio/document.xml", """
                <?xml version="1.0" encoding="utf-8"?>
                <VisioDocument xmlns="http://schemas.microsoft.com/office/visio/2012/main">
                  <DocumentSettings TopPage="0" DefaultTextStyle="0" DefaultLineStyle="0" DefaultFillStyle="0" DefaultGuideStyle="4" CustomSettingsFlag="KeepMe">
                    <GlueSettings>9</GlueSettings>
                    <SnapSettings>295</SnapSettings>
                    <SnapExtensions>34</SnapExtensions>
                    <SnapAngles />
                    <DynamicGridEnabled>1</DynamicGridEnabled>
                    <ProtectStyles>0</ProtectStyles>
                    <ProtectShapes>0</ProtectShapes>
                    <ProtectMasters>0</ProtectMasters>
                    <ProtectBkgnds>0</ProtectBkgnds>
                    <RelayoutAndRerouteUponOpen>1</RelayoutAndRerouteUponOpen>
                    <CustomSettingsElement Enabled="1" />
                  </DocumentSettings>
                  <Colors />
                  <FaceNames />
                  <StyleSheets>
                    <StyleSheet ID="0" Name="No Style" NameU="No Style">
                      <Cell N="EnableLineProps" V="1" />
                    </StyleSheet>
                  </StyleSheets>
                </VisioDocument>
                """);

            string originalFragments = ReadDocumentSettingsFragments(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalFragments, ReadDocumentSettingsFragments(savedPath));
        }

        [Fact]
        public void ColorsAndFaceNamesFragmentsArePreservedOnRoundTrip() {
            string filePath = CreateShapeDocument();

            RewriteEntry(filePath, "visio/document.xml", """
                <?xml version="1.0" encoding="utf-8"?>
                <VisioDocument xmlns="http://schemas.microsoft.com/office/visio/2012/main">
                  <DocumentSettings TopPage="0" DefaultTextStyle="0" DefaultLineStyle="0" DefaultFillStyle="0" DefaultGuideStyle="4">
                    <GlueSettings>9</GlueSettings>
                    <SnapSettings>295</SnapSettings>
                    <SnapExtensions>34</SnapExtensions>
                    <SnapAngles />
                    <DynamicGridEnabled>1</DynamicGridEnabled>
                    <ProtectStyles>0</ProtectStyles>
                    <ProtectShapes>0</ProtectShapes>
                    <ProtectMasters>0</ProtectMasters>
                    <ProtectBkgnds>0</ProtectBkgnds>
                  </DocumentSettings>
                  <Colors CustomPalette="KeepMe">
                    <ColorEntry IX="0" RGB="#112233" />
                    <ColorEntry IX="1" RGB="#445566" Name="AccentCustom" />
                  </Colors>
                  <FaceNames CustomFonts="KeepMeToo">
                    <FaceName ID="0" Name="Aptos" UnicodeRanges="0-255" CharSets="0" />
                    <FaceName ID="1" Name="Consolas" Panos="020B0609030504040204" />
                  </FaceNames>
                  <StyleSheets>
                    <StyleSheet ID="0" Name="No Style" NameU="No Style">
                      <Cell N="EnableLineProps" V="1" />
                    </StyleSheet>
                  </StyleSheets>
                </VisioDocument>
                """);

            string originalFragments = ReadColorsAndFaceNamesFragments(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalFragments, ReadColorsAndFaceNamesFragments(savedPath));
        }

        [Fact]
        public void PageContentsAttributesAndElementsArePreservedOnRoundTrip() {
            string filePath = CreateShapeDocument();

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                pageDoc.Root!.SetAttributeValue("CustomPageContent", "KeepMe");
                pageDoc.Root.AddFirst(new XElement(ns + "Reviewers", new XAttribute("Status", "Open")));
            });

            string originalFragments = ReadFirstPageContentFragments(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalFragments, ReadFirstPageContentFragments(savedPath));
        }

        [Fact]
        public void PageShapesAndConnectsContainerFragmentsArePreservedOnRoundTrip() {
            string filePath = CreateConnectorDocument(ConnectorKind.Straight);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement shapes = pageDoc.Root!.Element(ns + "Shapes")!;
                shapes.SetAttributeValue("ContainerFlag", "KeepShapes");
                shapes.AddFirst(new XElement(ns + "ShapeCatalog", new XAttribute("Version", "1")));

                XElement connects = pageDoc.Root.Element(ns + "Connects")!;
                connects.SetAttributeValue("RoutingMode", "KeepConnects");
                connects.AddFirst(new XElement(ns + "ConnectCatalog", new XAttribute("Version", "2")));
            });

            string originalFragments = ReadFirstPageContainerFragments(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalFragments, ReadFirstPageContainerFragments(savedPath));
        }

        [Fact]
        public void ShapesChildOrderIsPreservedOnRoundTrip() {
            string filePath = CreateConnectorDocument(ConnectorKind.Straight);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement shapes = pageDoc.Root!.Element(ns + "Shapes")!;
                XElement[] originalShapes = shapes.Elements(ns + "Shape").ToArray();
                XElement firstShape = originalShapes[0];
                XElement secondShape = originalShapes[1];
                XElement connectorShape = originalShapes[2];
                XElement sidecarElement = new(ns + "ShapesMeta",
                    new XAttribute("KeepOrder", "1"));

                shapes.RemoveNodes();
                shapes.Add(connectorShape, sidecarElement, secondShape, firstShape);
            });

            string originalOrder = ReadShapesChildOrder(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalOrder, ReadShapesChildOrder(savedPath));
        }

        [Fact]
        public void RuntimeShapeOrderOverridesPreservedShapesOrderOnSave() {
            string filePath = CreateConnectorDocument(ConnectorKind.Straight);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement shapes = pageDoc.Root!.Element(ns + "Shapes")!;
                XElement[] originalShapes = shapes.Elements(ns + "Shape").ToArray();
                XElement firstShape = originalShapes[0];
                XElement secondShape = originalShapes[1];
                XElement connectorShape = originalShapes[2];

                shapes.RemoveNodes();
                shapes.Add(connectorShape, secondShape, firstShape);
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage page = loaded.Pages[0];
            VisioShape start = page.Shapes[0];

            page.Shapes.RemoveAt(0);
            page.Shapes.Add(start);

            string expectedOrder = "Shape:Connector|" + string.Join("|", page.Shapes.Select(shape => $"Shape:{shape.NameU}"));
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(expectedOrder, ReadShapesChildOrder(savedPath));
        }

        [Fact]
        public void ConnectorShapeChildOrderIsPreservedOnRoundTrip() {
            string filePath = CreateConnectorDocument(ConnectorKind.Straight);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement connectorShape = GetConnectorShape(pageDoc, ns);
                XElement xForm = connectorShape.Element(ns + "XForm1D")!;
                XElement lineWeight = connectorShape.Elements(ns + "Cell")
                    .First(cell => (string?)cell.Attribute("N") == "LineWeight");
                XElement geometry = connectorShape.Elements(ns + "Section")
                    .First(section => (string?)section.Attribute("N") == "Geometry");
                XElement sidecarElement = new(ns + "ShapeMeta",
                    new XAttribute("KeepOrder", "1"));

                XElement[] remainingChildren = connectorShape.Elements()
                    .Where(element => element != xForm && element != lineWeight && element != geometry)
                    .ToArray();

                connectorShape.RemoveNodes();
                connectorShape.Add(lineWeight, sidecarElement, geometry, xForm);
                foreach (XElement remainingChild in remainingChildren) {
                    connectorShape.Add(remainingChild);
                }
            });

            string originalOrder = ReadConnectorShapeChildOrder(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalOrder, ReadConnectorShapeChildOrder(savedPath));
        }

        [Fact]
        public void AdditionalConnectRowsArePreservedOnRoundTrip() {
            string filePath = CreateConnectorDocument(ConnectorKind.Straight);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement connects = pageDoc.Root!.Element(ns + "Connects")!;
                XElement beginConnect = connects.Elements(ns + "Connect")
                    .First(connect => (string?)connect.Attribute("FromCell") == "BeginX");
                connects.AddFirst(new XElement(ns + "Connect",
                    new XAttribute("FromSheet", beginConnect.Attribute("FromSheet")!.Value),
                    new XAttribute("FromCell", "PinX"),
                    new XAttribute("ToSheet", beginConnect.Attribute("ToSheet")!.Value),
                    new XAttribute("ToCell", "Connections.X1"),
                    new XAttribute("CustomConnectFlag", "KeepMe")));
            });

            string originalFragments = ReadAdditionalConnectRows(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalFragments, ReadAdditionalConnectRows(savedPath));
        }

        [Fact]
        public void ConnectRowOrderIsPreservedOnRoundTrip() {
            string filePath = CreateConnectorDocument(ConnectorKind.Straight);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement connects = pageDoc.Root!.Element(ns + "Connects")!;
                XElement[] originalRows = connects.Elements(ns + "Connect").ToArray();
                XElement beginRow = originalRows.First(connect => (string?)connect.Attribute("FromCell") == "BeginX");
                XElement endRow = originalRows.First(connect => (string?)connect.Attribute("FromCell") == "EndX");
                XElement extraRow = new(ns + "Connect",
                    new XAttribute("FromSheet", beginRow.Attribute("FromSheet")!.Value),
                    new XAttribute("FromCell", "PinX"),
                    new XAttribute("ToSheet", beginRow.Attribute("ToSheet")!.Value),
                    new XAttribute("ToCell", "Connections.X1"),
                    new XAttribute("CustomConnectFlag", "KeepOrder"));

                connects.RemoveNodes();
                connects.Add(endRow, extraRow, beginRow);
            });

            string originalOrder = ReadConnectRowOrder(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalOrder, ReadConnectRowOrder(savedPath));
        }

        [Fact]
        public void ConnectsChildOrderIsPreservedOnRoundTrip() {
            string filePath = CreateConnectorDocument(ConnectorKind.Straight);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement connects = pageDoc.Root!.Element(ns + "Connects")!;
                XElement[] originalRows = connects.Elements(ns + "Connect").ToArray();
                XElement beginRow = originalRows.First(connect => (string?)connect.Attribute("FromCell") == "BeginX");
                XElement endRow = originalRows.First(connect => (string?)connect.Attribute("FromCell") == "EndX");
                XElement sidecarElement = new(ns + "ConnectionMeta",
                    new XAttribute("KeepOrder", "1"));

                connects.RemoveNodes();
                connects.Add(beginRow, sidecarElement, endRow);
            });

            string originalOrder = ReadConnectChildOrder(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalOrder, ReadConnectChildOrder(savedPath));
        }

        [Fact]
        public void NamespacedConnectAttributesArePreservedOnRoundTrip() {
            string filePath = CreateConnectorDocument(ConnectorKind.Straight);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XNamespace cx = "urn:officeimo:connect-meta";
                XElement beginConnect = pageDoc.Root!
                    .Element(ns + "Connects")!
                    .Elements(ns + "Connect")
                    .First(connect => (string?)connect.Attribute("FromCell") == "BeginX");
                beginConnect.SetAttributeValue(cx + "tag", "PreserveMe");
            });

            string originalAttributes = ReadNamespacedConnectAttributes(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalAttributes, ReadNamespacedConnectAttributes(savedPath));
        }

        [Fact]
        public void ConnectAttributeOrderIsPreservedOnRoundTrip() {
            string filePath = CreateConnectorDocument(ConnectorKind.Straight);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XNamespace cx = "urn:officeimo:connect-meta";
                XElement beginConnect = pageDoc.Root!
                    .Element(ns + "Connects")!
                    .Elements(ns + "Connect")
                    .First(connect => (string?)connect.Attribute("FromCell") == "BeginX");

                XAttribute fromSheet = beginConnect.Attribute("FromSheet")!;
                XAttribute fromCell = beginConnect.Attribute("FromCell")!;
                XAttribute toSheet = beginConnect.Attribute("ToSheet")!;
                XAttribute toCell = beginConnect.Attribute("ToCell")!;
                beginConnect.RemoveAttributes();
                beginConnect.Add(
                    new XAttribute(cx + "tag", "Order"),
                    toCell,
                    fromSheet,
                    toSheet,
                    fromCell);
            });

            string originalOrder = ReadConnectAttributeOrder(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalOrder, ReadConnectAttributeOrder(savedPath));
        }

        [Fact]
        public void StyleSheetsFragmentsArePreservedOnRoundTrip() {
            string filePath = CreateShapeDocument();

            RewriteEntry(filePath, "visio/document.xml", """
                <?xml version="1.0" encoding="utf-8"?>
                <VisioDocument xmlns="http://schemas.microsoft.com/office/visio/2012/main">
                  <DocumentSettings TopPage="0" DefaultTextStyle="0" DefaultLineStyle="0" DefaultFillStyle="0" DefaultGuideStyle="4">
                    <GlueSettings>9</GlueSettings>
                    <SnapSettings>295</SnapSettings>
                    <SnapExtensions>34</SnapExtensions>
                    <SnapAngles />
                    <DynamicGridEnabled>1</DynamicGridEnabled>
                    <ProtectStyles>0</ProtectStyles>
                    <ProtectShapes>0</ProtectShapes>
                    <ProtectMasters>0</ProtectMasters>
                    <ProtectBkgnds>0</ProtectBkgnds>
                  </DocumentSettings>
                  <Colors />
                  <FaceNames />
                  <StyleSheets CustomStyleRoot="KeepMe">
                    <StyleCatalog Version="2" />
                    <StyleSheet ID="0" Name="No Style" NameU="No Style" CustomAttr0="Keep0">
                      <Cell N="EnableLineProps" V="1" />
                      <Cell N="EnableFillProps" V="1" />
                      <Cell N="EnableTextProps" V="1" />
                      <Cell N="LineWeight" V="0.01041666666666667" />
                      <Cell N="LineColor" V="0" />
                      <Cell N="LinePattern" V="1" />
                      <Cell N="FillForegnd" V="1" />
                      <Cell N="FillPattern" V="1" />
                      <Section N="User"><Row N="Style0Flag"><Cell N="Value" V="TRUE" /></Row></Section>
                    </StyleSheet>
                    <StyleSheet ID="1" Name="Normal" NameU="Normal" BasedOn="0" LineStyle="0" FillStyle="0" TextStyle="0" CustomAttr1="Keep1">
                      <Cell N="LinePattern" V="1" />
                      <Cell N="LineColor" V="#000000" />
                      <Cell N="FillPattern" V="1" />
                      <Cell N="FillForegnd" V="#FFFFFF" />
                      <Cell N="LineCap" V="2" />
                    </StyleSheet>
                    <StyleSheet ID="2" Name="Connector" NameU="Connector" BasedOn="1" LineStyle="0" FillStyle="0" TextStyle="0" CustomAttr2="Keep2">
                      <Cell N="EndArrow" V="0" />
                      <Cell N="LineColor" V="#005A9C" />
                    </StyleSheet>
                    <StyleSheet ID="7" Name="Visio Authored" NameU="Visio Authored">
                      <Cell N="LineColor" V="#FF0000" />
                    </StyleSheet>
                  </StyleSheets>
                </VisioDocument>
                """);

            string originalFragments = ReadStyleSheetsFragments(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalFragments, ReadStyleSheetsFragments(savedPath));
        }

        [Fact]
        public void GeneratedStyleSheetZeroDoesNotDuplicateBaseStyleAttributesOnRoundTrip() {
            string filePath = CreateShapeDocument();

            RewriteEntry(filePath, "visio/document.xml", """
                <?xml version="1.0" encoding="utf-8"?>
                <VisioDocument xmlns="http://schemas.microsoft.com/office/visio/2012/main">
                  <DocumentSettings TopPage="0" DefaultTextStyle="0" DefaultLineStyle="0" DefaultFillStyle="0" DefaultGuideStyle="4">
                    <GlueSettings>9</GlueSettings>
                    <SnapSettings>295</SnapSettings>
                    <SnapExtensions>34</SnapExtensions>
                    <SnapAngles />
                    <DynamicGridEnabled>1</DynamicGridEnabled>
                    <ProtectStyles>0</ProtectStyles>
                    <ProtectShapes>0</ProtectShapes>
                    <ProtectMasters>0</ProtectMasters>
                    <ProtectBkgnds>0</ProtectBkgnds>
                  </DocumentSettings>
                  <Colors />
                  <FaceNames />
                  <StyleSheets>
                    <StyleSheet ID="0" Name="No Style" NameU="No Style" BasedOn="0" LineStyle="0" FillStyle="0" TextStyle="0" CustomAttr0="Keep0">
                      <Cell N="EnableLineProps" V="1" />
                      <Cell N="EnableFillProps" V="1" />
                      <Cell N="EnableTextProps" V="1" />
                      <Cell N="LineWeight" V="0.01041666666666667" />
                      <Cell N="LineColor" V="0" />
                      <Cell N="LinePattern" V="1" />
                      <Cell N="FillForegnd" V="1" />
                      <Cell N="FillPattern" V="1" />
                    </StyleSheet>
                  </StyleSheets>
                </VisioDocument>
                """);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            using ZipArchive archive = ZipFile.OpenRead(savedPath);
            using Stream stream = archive.GetEntry("visio/document.xml")!.Open();
            XDocument documentXml = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement styleSheet = documentXml.Root!
                .Element(ns + "StyleSheets")!
                .Elements(ns + "StyleSheet")
                .First(style => (string?)style.Attribute("ID") == "0");

            Assert.Null(styleSheet.Attribute("BasedOn"));
            Assert.Null(styleSheet.Attribute("LineStyle"));
            Assert.Null(styleSheet.Attribute("FillStyle"));
            Assert.Null(styleSheet.Attribute("TextStyle"));
            Assert.Equal("Keep0", (string?)styleSheet.Attribute("CustomAttr0"));
            Assert.Single(styleSheet.Attributes(), attribute =>
                string.Equals(attribute.Name.LocalName, "CustomAttr0", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void CustomConnectorGeometryIsPreservedOnRoundTrip() {
            string filePath = CreateConnectorDocument(ConnectorKind.Straight);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement connectorShape = GetConnectorShape(pageDoc, ns);
                XElement geometry = connectorShape.Elements(ns + "Section").First(section => (string?)section.Attribute("N") == "Geometry");
                geometry.RemoveNodes();
                geometry.Add(
                    new XElement(ns + "Row",
                        new XAttribute("T", "Geometry"),
                        new XElement(ns + "Cell", new XAttribute("N", "NoFill"), new XAttribute("V", "0"))),
                    CreateRow(ns, "MoveTo", 1, 1),
                    CreateRow(ns, "LineTo", 2, 4),
                    CreateRow(ns, "LineTo", 4, 3),
                    CreateRow(ns, "LineTo", 5, 2));
            });

            string originalGeometry = ReadConnectorGeometry(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.True(XNode.DeepEquals(
                XElement.Parse(originalGeometry),
                XElement.Parse(ReadConnectorGeometry(savedPath))));
        }

        [Fact]
        public void ConnectorLayoutCellsAndControlSectionsArePreservedOnRoundTrip() {
            string filePath = CreateConnectorDocument(ConnectorKind.Straight);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement connectorShape = GetConnectorShape(pageDoc, ns);
                connectorShape.Add(
                    new XElement(ns + "Cell",
                        new XAttribute("N", "TxtPinX"),
                        new XAttribute("V", "2.25"),
                        new XAttribute("F", "Width*0.5")),
                    new XElement(ns + "Cell",
                        new XAttribute("N", "TxtPinY"),
                        new XAttribute("V", "-0.75"),
                        new XAttribute("F", "Controls.Row_1.Y")));
                connectorShape.Add(
                    new XElement(ns + "Section",
                        new XAttribute("N", "Control"),
                        new XAttribute("IX", "0"),
                        new XElement(ns + "Row",
                            new XAttribute("N", "Row_1"),
                            new XAttribute("IX", "0"),
                            new XAttribute("T", "Control"),
                            new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "1.5")),
                            new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "-0.75")),
                            new XElement(ns + "Cell", new XAttribute("N", "XDyn"), new XAttribute("V", "1.5")),
                            new XElement(ns + "Cell", new XAttribute("N", "YDyn"), new XAttribute("V", "-0.75")))));
            });

            string originalFragments = ReadConnectorLayoutFragments(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalFragments, ReadConnectorLayoutFragments(savedPath));
        }

        [Fact]
        public void CustomShapeGeometryIsPreservedOnRoundTrip() {
            string filePath = CreateShapeDocument();

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement shape = GetFirstShape(pageDoc, ns);
                shape.SetAttributeValue("NameU", "CustomChevron");

                XElement geometry = shape.Elements(ns + "Section").First(section => (string?)section.Attribute("N") == "Geometry");
                geometry.RemoveNodes();
                geometry.Add(
                    new XElement(ns + "Row",
                        new XAttribute("T", "Geometry"),
                        new XElement(ns + "Cell", new XAttribute("N", "NoFill"), new XAttribute("V", "0")),
                        new XElement(ns + "Cell", new XAttribute("N", "NoLine"), new XAttribute("V", "0"))),
                    CreateRow(ns, "MoveTo", 0, 0),
                    CreateRow(ns, "LineTo", 1.5, 0),
                    CreateRow(ns, "LineTo", 2, 0.5),
                    CreateRow(ns, "LineTo", 1.5, 1),
                    CreateRow(ns, "LineTo", 0, 1),
                    CreateRow(ns, "LineTo", 0.5, 0.5),
                    CreateRow(ns, "LineTo", 0, 0));
            });

            string originalGeometry = ReadFirstShapeGeometry(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.True(XNode.DeepEquals(
                XElement.Parse(originalGeometry),
                XElement.Parse(ReadFirstShapeGeometry(savedPath))));
        }

        [Fact]
        public void ShapeLayoutCellsAndUserSectionsArePreservedOnRoundTrip() {
            string filePath = CreateShapeDocument();

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement shape = GetFirstShape(pageDoc, ns);
                shape.Add(
                    new XElement(ns + "Cell",
                        new XAttribute("N", "TxtPinX"),
                        new XAttribute("V", "1.125"),
                        new XAttribute("F", "Width*0.5")),
                    new XElement(ns + "Cell",
                        new XAttribute("N", "TxtPinY"),
                        new XAttribute("V", "0.75"),
                        new XAttribute("F", "Height*0.75")));
                shape.Add(
                    new XElement(ns + "Section",
                        new XAttribute("N", "User"),
                        new XAttribute("IX", "0"),
                        new XElement(ns + "Row",
                            new XAttribute("N", "CustomFlag"),
                            new XAttribute("IX", "0"),
                            new XAttribute("T", "User"),
                            new XElement(ns + "Cell", new XAttribute("N", "Value"), new XAttribute("V", "TRUE")))));
            });

            string originalFragments = ReadFirstShapeLayoutFragments(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalFragments, ReadFirstShapeLayoutFragments(savedPath));
        }

        [Fact]
        public void ShapeRichTextMarkupIsPreservedOnRoundTripWhenTextIsUnchanged() {
            string filePath = CreateShapeDocument();

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement shape = GetFirstShape(pageDoc, ns);
                shape.Element(ns + "Text")!.ReplaceWith(
                    new XElement(ns + "Text",
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        new XElement(ns + "cp", new XAttribute("IX", "0")),
                        "Alpha ",
                        new XElement(ns + "pp", new XAttribute("IX", "0")),
                        "Beta"));
            });

            string originalText = ReadFirstShapeTextElement(filePath).ToString(SaveOptions.DisableFormatting);

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal("Alpha Beta", Assert.Single(loaded.Pages[0].Shapes).Text);

            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalText, ReadFirstShapeTextElement(savedPath).ToString(SaveOptions.DisableFormatting));
        }

        [Fact]
        public void ConnectorRichTextMarkupIsPreservedOnRoundTripWhenLabelIsUnchanged() {
            string filePath = CreateConnectorDocument(ConnectorKind.Straight);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement connectorShape = GetConnectorShape(pageDoc, ns);
                connectorShape.Add(
                    new XElement(ns + "Text",
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        new XElement(ns + "cp", new XAttribute("IX", "1")),
                        "Edge ",
                        new XElement(ns + "tp", new XAttribute("IX", "0")),
                        "Label"));
            });

            string originalText = ReadConnectorTextElement(filePath).ToString(SaveOptions.DisableFormatting);

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal("Edge Label", Assert.Single(loaded.Pages[0].Connectors).Label);

            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalText, ReadConnectorTextElement(savedPath).ToString(SaveOptions.DisableFormatting));
        }

        [Fact]
        public void UpdatingShapeTextReplacesPreservedRichTextMarkup() {
            string filePath = CreateShapeDocument();

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement shape = GetFirstShape(pageDoc, ns);
                shape.Element(ns + "Text")!.ReplaceWith(
                    new XElement(ns + "Text",
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        new XElement(ns + "cp", new XAttribute("IX", "0")),
                        "Original"));
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Single(loaded.Pages[0].Shapes).Text = "Updated";

            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            XElement savedText = ReadFirstShapeTextElement(savedPath);
            Assert.Equal("Updated", savedText.Value);
            Assert.Empty(savedText.Elements());
        }

        [Fact]
        public void ShapeDataRowsPreserveMetadataOnRoundTrip() {
            string filePath = CreateShapeDocument();

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement shape = GetFirstShape(pageDoc, ns);
                shape.Add(
                    new XElement(ns + "Section",
                        new XAttribute("N", "Prop"),
                        new XElement(ns + "Row",
                            new XAttribute("N", "Status"),
                            new XAttribute("IX", "0"),
                            new XAttribute("T", "Prop"),
                            new XElement(ns + "Cell", new XAttribute("N", "Label"), new XAttribute("V", "Ticket status")),
                            new XElement(ns + "Cell", new XAttribute("N", "Prompt"), new XAttribute("V", "Choose status")),
                            new XElement(ns + "Cell", new XAttribute("N", "Format"), new XAttribute("V", "\"Open;Closed\"")),
                            new XElement(ns + "Cell", new XAttribute("N", "SortKey"), new XAttribute("V", "10")),
                            new XElement(ns + "Cell", new XAttribute("N", "Value"), new XAttribute("V", "Open")))));
            });

            string originalRow = ReadFirstShapeDataRow(filePath, "Status").ToString(SaveOptions.DisableFormatting);

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal("Open", Assert.Single(loaded.Pages[0].Shapes).Data["Status"]);

            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalRow, ReadFirstShapeDataRow(savedPath, "Status").ToString(SaveOptions.DisableFormatting));
        }

        [Fact]
        public void UpdatingShapeDataValueKeepsPreservedRowMetadata() {
            string filePath = CreateShapeDocument();

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement shape = GetFirstShape(pageDoc, ns);
                shape.Add(
                    new XElement(ns + "Section",
                        new XAttribute("N", "Prop"),
                        new XElement(ns + "Row",
                            new XAttribute("N", "Status"),
                            new XAttribute("IX", "0"),
                            new XAttribute("T", "Prop"),
                            new XElement(ns + "Cell", new XAttribute("N", "Label"), new XAttribute("V", "Ticket status")),
                            new XElement(ns + "Cell", new XAttribute("N", "Prompt"), new XAttribute("V", "Choose status")),
                            new XElement(ns + "Cell", new XAttribute("N", "Type"), new XAttribute("V", "0")),
                            new XElement(ns + "Cell", new XAttribute("N", "Value"), new XAttribute("V", "Open")))));
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape shape = Assert.Single(loaded.Pages[0].Shapes);
            shape.Data["Status"] = "Closed";

            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            XElement savedRow = ReadFirstShapeDataRow(savedPath, "Status");
            Assert.Equal("Closed", savedRow.Elements(savedRow.Name.Namespace + "Cell")
                .Single(cell => (string?)cell.Attribute("N") == "Value")
                .Attribute("V")?.Value);
            Assert.Equal("Ticket status", savedRow.Elements(savedRow.Name.Namespace + "Cell")
                .Single(cell => (string?)cell.Attribute("N") == "Label")
                .Attribute("V")?.Value);
            Assert.Equal("Choose status", savedRow.Elements(savedRow.Name.Namespace + "Cell")
                .Single(cell => (string?)cell.Attribute("N") == "Prompt")
                .Attribute("V")?.Value);
            Assert.Equal("0", savedRow.Elements(savedRow.Name.Namespace + "Cell")
                .Single(cell => (string?)cell.Attribute("N") == "Type")
                .Attribute("V")?.Value);
        }

        [Fact]
        public void UpdatingShapeDataValueClearsPreservedValueFormula() {
            string filePath = CreateShapeDocument();

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement shape = GetFirstShape(pageDoc, ns);
                shape.Add(
                    new XElement(ns + "Section",
                        new XAttribute("N", "Prop"),
                        new XElement(ns + "Row",
                            new XAttribute("N", "Status"),
                            new XAttribute("IX", "0"),
                            new XAttribute("T", "Prop"),
                            new XElement(ns + "Cell", new XAttribute("N", "Value"), new XAttribute("V", "Open"), new XAttribute("F", "\"Open\"")))));
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape shape = Assert.Single(loaded.Pages[0].Shapes);
            shape.Data["Status"] = "Closed";

            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            XElement valueCell = ReadFirstShapeDataRow(savedPath, "Status")
                .Elements(XName.Get("Cell", "http://schemas.microsoft.com/office/visio/2012/main"))
                .Single(cell => (string?)cell.Attribute("N") == "Value");
            Assert.Equal("Closed", (string?)valueCell.Attribute("V"));
            Assert.Null(valueCell.Attribute("F"));
        }

        [Fact]
        public void PageSheetCellsAndSectionsArePreservedOnRoundTrip() {
            string filePath = CreateShapeDocument();

            RewritePages(filePath, pagesDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement pageSheet = GetFirstPageSheet(pagesDoc, ns);
                pageSheet.Add(
                    new XElement(ns + "Cell",
                        new XAttribute("N", "XRulerDensity"),
                        new XAttribute("V", "8")),
                    new XElement(ns + "Cell",
                        new XAttribute("N", "YRulerDensity"),
                        new XAttribute("V", "4")));
                pageSheet.Add(
                    new XElement(ns + "Section",
                        new XAttribute("N", "User"),
                        new XElement(ns + "Row",
                            new XAttribute("N", "CustomPageFlag"),
                            new XAttribute("IX", "0"),
                            new XAttribute("T", "User"),
                            new XElement(ns + "Cell", new XAttribute("N", "Value"), new XAttribute("V", "TRUE")))));
            });

            string originalFragments = ReadFirstPageSheetFragments(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalFragments, ReadFirstPageSheetFragments(savedPath));
        }

        [Fact]
        public void GeneratedPageSheetCellsAreNotDuplicatedOnRoundTrip() {
            string filePath = CreateShapeDocument();

            RewritePages(filePath, pagesDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement pageSheet = GetFirstPageSheet(pagesDoc, ns);
                pageSheet.Add(
                    new XElement(ns + "Cell", new XAttribute("N", "DrawingResizeType"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "PageShapeSplit"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "UIVisibility"), new XAttribute("V", "0")));
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            using ZipArchive archive = ZipFile.OpenRead(savedPath);
            using Stream stream = archive.GetEntry("visio/pages/pages.xml")!.Open();
            XDocument pagesDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement pageSheet = GetFirstPageSheet(pagesDoc, ns);
            Assert.Equal(1, pageSheet.Elements(ns + "Cell").Count(cell => (string?)cell.Attribute("N") == "DrawingResizeType"));
            Assert.Equal(1, pageSheet.Elements(ns + "Cell").Count(cell => (string?)cell.Attribute("N") == "PageShapeSplit"));
            Assert.Equal(1, pageSheet.Elements(ns + "Cell").Count(cell => (string?)cell.Attribute("N") == "UIVisibility"));
        }

        [Fact]
        public void PageAttributesArePreservedOnRoundTrip() {
            string filePath = CreateShapeDocument();

            RewritePages(filePath, pagesDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement page = pagesDoc.Root!.Elements(ns + "Page").First();
                page.SetAttributeValue("ReviewerID", "7");
                page.SetAttributeValue("AssociatedPage", "2");
            });

            string originalAttributes = ReadFirstPageAttributes(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalAttributes, ReadFirstPageAttributes(savedPath));
        }

        [Fact]
        public void MasterAttributesAndPageSheetFragmentsArePreservedOnRoundTrip() {
            string filePath = CreateMasterBackedShapeDocument();

            RewriteMasters(filePath, mastersDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement master = mastersDoc.Root!.Elements(ns + "Master").First();
                master.SetAttributeValue("PageType", "Foreground");
                master.SetAttributeValue("IconFlags", "3");

                XElement pageSheet = master.Element(ns + "PageSheet")!;
                pageSheet.SetAttributeValue("UniqueID", "{11111111-1111-1111-1111-111111111111}");
                pageSheet.Add(
                    new XElement(ns + "Cell",
                        new XAttribute("N", "XRulerDensity"),
                        new XAttribute("V", "12")));
                pageSheet.Add(
                    new XElement(ns + "Section",
                        new XAttribute("N", "User"),
                        new XElement(ns + "Row",
                            new XAttribute("N", "CustomMasterFlag"),
                            new XAttribute("IX", "0"),
                            new XAttribute("T", "User"),
                            new XElement(ns + "Cell", new XAttribute("N", "Value"), new XAttribute("V", "TRUE")))));
            });

            string originalAttributes = ReadFirstMasterAttributes(filePath);
            string originalPageSheet = ReadFirstMasterPageSheetFragments(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalAttributes, ReadFirstMasterAttributes(savedPath));
            Assert.Equal(originalPageSheet, ReadFirstMasterPageSheetFragments(savedPath));
        }

        [Fact]
        public void MasterCatalogChildElementsArePreservedOnRoundTrip() {
            string filePath = CreateMasterBackedShapeDocument();

            RewriteMasters(filePath, mastersDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement master = mastersDoc.Root!.Elements(ns + "Master").First();
                master.AddFirst(new XElement(ns + "CustomMetadata",
                    new XAttribute("Flag", "1"),
                    new XElement(ns + "Value", "RetainMe")));
            });

            string originalFragments = ReadFirstMasterChildElements(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalFragments, ReadFirstMasterChildElements(savedPath));
        }

        [Fact]
        public void MastersRootAttributesAndElementsArePreservedOnRoundTrip() {
            string filePath = CreateMasterBackedShapeDocument();

            RewriteMasters(filePath, mastersDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement masters = mastersDoc.Root!;
                masters.SetAttributeValue("IconCacheVersion", "5");
                masters.AddFirst(new XElement(ns + "CustomCatalogMetadata",
                    new XAttribute("Scope", "Global"),
                    new XElement(ns + "Value", "RetainMe")));
            });

            string originalFragments = ReadMastersRootFragments(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalFragments, ReadMastersRootFragments(savedPath));
        }

        [Fact]
        public void MastersRootMetadataSurvivesWhenNewGeneratedMasterSortsBeforeLoadedMaster() {
            string filePath = CreateMasterBackedShapeDocument();

            RewriteMasters(filePath, mastersDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement masters = mastersDoc.Root!;
                masters.SetAttributeValue("IconCacheVersion", "5");
                masters.AddFirst(new XElement(ns + "CustomCatalogMetadata",
                    new XAttribute("Scope", "Global"),
                    new XElement(ns + "Value", "RetainMe")));
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            loaded.UseMastersByDefault = true;
            loaded.Pages[0].Shapes.Insert(0, new VisioShape("2", 4, 1, 2, 1, "Diamond") { NameU = "Diamond" });

            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(ReadMastersRootFragments(filePath), ReadMastersRootFragments(savedPath));
        }

        [Fact]
        public void MasterContentsAdditionalShapesAndRootElementsArePreservedOnRoundTrip() {
            string filePath = CreateMasterBackedShapeDocument();

            RewriteEntry(filePath, "visio/masters/master1.xml", """
                <?xml version="1.0" encoding="utf-8"?>
                <MasterContents xmlns="http://schemas.microsoft.com/office/visio/2012/main" xml:space="preserve">
                  <Shapes>
                    <Shape ID="1" Name="Rectangle" NameU="Rectangle" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">
                      <XForm>
                        <PinX>1</PinX>
                        <PinY>0.5</PinY>
                        <Width>2</Width>
                        <Height>1</Height>
                        <LocPinX>1</LocPinX>
                        <LocPinY>0.5</LocPinY>
                        <Angle>0</Angle>
                      </XForm>
                    </Shape>
                    <Shape ID="2" Name="Badge" NameU="Badge" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">
                      <Cell N="PinX" V="0.5" />
                      <Cell N="PinY" V="0.5" />
                      <Cell N="Width" V="0.5" />
                      <Cell N="Height" V="0.5" />
                      <Cell N="LocPinX" V="0.25" />
                      <Cell N="LocPinY" V="0.25" />
                      <Text>Extra</Text>
                    </Shape>
                  </Shapes>
                  <Connects />
                </MasterContents>
                """);

            string originalFragments = ReadFirstMasterContentFragments(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalFragments, ReadFirstMasterContentFragments(savedPath));
        }

        [Fact]
        public void MasterContentsAttributesAndShapesContainerAttributesArePreservedOnRoundTrip() {
            string filePath = CreateMasterBackedShapeDocument();

            RewriteEntry(filePath, "visio/masters/master1.xml", """
                <?xml version="1.0" encoding="utf-8"?>
                <MasterContents xmlns="http://schemas.microsoft.com/office/visio/2012/main" xml:space="preserve" CustomRootFlag="RetainMe">
                  <Shapes ForeignData="KeepThis">
                    <Shape ID="1" Name="Rectangle" NameU="Rectangle" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">
                      <XForm>
                        <PinX>1</PinX>
                        <PinY>0.5</PinY>
                        <Width>2</Width>
                        <Height>1</Height>
                        <LocPinX>1</LocPinX>
                        <LocPinY>0.5</LocPinY>
                        <Angle>0</Angle>
                      </XForm>
                    </Shape>
                  </Shapes>
                </MasterContents>
                """);

            string originalAttributes = ReadFirstMasterContentAttributes(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Equal(originalAttributes, ReadFirstMasterContentAttributes(savedPath));
        }

        [Fact]
        public void LoadInfersNameUFromResolvedMasterWhenPageShapeOmitsIt() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            document.UseMastersByDefault = true;
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 2, 2, 2, 1, "Rectangle") { NameU = "Rectangle" });
            document.Save();

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement shape = GetFirstShape(pageDoc, ns);
                shape.Attribute("NameU")?.Remove();
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape shape = Assert.Single(loaded.Pages[0].Shapes);

            Assert.Equal("Rectangle", shape.NameU);
            Assert.Equal("Rectangle", shape.Master?.NameU);
        }

        private static string CreateConnectorDocument(ConnectorKind kind) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");
            VisioShape start = new("1", 1, 1, 1, 1, "Start");
            VisioShape end = new("2", 4, 2, 1, 1, "End");
            page.Shapes.Add(start);
            page.Shapes.Add(end);
            page.Connectors.Add(new VisioConnector(start, end) { Kind = kind });
            document.Save();
            return filePath;
        }

        private static string CreateShapeDocument() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 2, 2, 2, 1, "Shape"));
            document.Save();
            return filePath;
        }

        private static string CreateThemedDocument() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.Theme = new VisioTheme { Name = "Office Theme" };
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, string.Empty));
            document.Save();
            return filePath;
        }

        private static string CreateMasterBackedShapeDocument() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.UseMastersByDefault = true;
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Rectangle") { NameU = "Rectangle" });
            document.Save();
            return filePath;
        }

        private static XElement GetConnectorShape(XDocument pageDoc, XNamespace ns) {
            return pageDoc.Root!
                .Element(ns + "Shapes")!
                .Elements(ns + "Shape")
                .Last();
        }

        private static XElement GetFirstShape(XDocument pageDoc, XNamespace ns) {
            return pageDoc.Root!
                .Element(ns + "Shapes")!
                .Elements(ns + "Shape")
                .First();
        }

        private static XElement CreateRow(XNamespace ns, string type, double x, double y) {
            return new XElement(ns + "Row",
                new XAttribute("T", type),
                new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", x.ToString(System.Globalization.CultureInfo.InvariantCulture))),
                new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", y.ToString(System.Globalization.CultureInfo.InvariantCulture))));
        }

        private static void RewritePage(string vsdxPath, Action<XDocument> transform) {
            using FileStream stream = File.Open(vsdxPath, FileMode.Open, FileAccess.ReadWrite);
            using ZipArchive archive = new(stream, ZipArchiveMode.Update);
            ZipArchiveEntry pageEntry = archive.GetEntry("visio/pages/page1.xml")!;
            XDocument pageDoc;
            using (Stream pageStream = pageEntry.Open()) {
                pageDoc = XDocument.Load(pageStream);
            }

            transform(pageDoc);
            pageEntry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry("visio/pages/page1.xml");
            using Stream replacementStream = replacement.Open();
            using StreamWriter writer = new(replacementStream, new UTF8Encoding(false));
            writer.Write(pageDoc.Declaration + Environment.NewLine + pageDoc.ToString(SaveOptions.DisableFormatting));
        }

        private static void RewritePages(string vsdxPath, Action<XDocument> transform) {
            using FileStream stream = File.Open(vsdxPath, FileMode.Open, FileAccess.ReadWrite);
            using ZipArchive archive = new(stream, ZipArchiveMode.Update);
            ZipArchiveEntry pagesEntry = archive.GetEntry("visio/pages/pages.xml")!;
            XDocument pagesDoc;
            using (Stream pagesStream = pagesEntry.Open()) {
                pagesDoc = XDocument.Load(pagesStream);
            }

            transform(pagesDoc);
            pagesEntry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry("visio/pages/pages.xml");
            using Stream replacementStream = replacement.Open();
            using StreamWriter writer = new(replacementStream, new UTF8Encoding(false));
            writer.Write(pagesDoc.Declaration + Environment.NewLine + pagesDoc.ToString(SaveOptions.DisableFormatting));
        }

        private static void RewriteMasters(string vsdxPath, Action<XDocument> transform) {
            using FileStream stream = File.Open(vsdxPath, FileMode.Open, FileAccess.ReadWrite);
            using ZipArchive archive = new(stream, ZipArchiveMode.Update);
            ZipArchiveEntry mastersEntry = archive.GetEntry("visio/masters/masters.xml")!;
            XDocument mastersDoc;
            using (Stream mastersStream = mastersEntry.Open()) {
                mastersDoc = XDocument.Load(mastersStream);
            }

            transform(mastersDoc);
            mastersEntry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry("visio/masters/masters.xml");
            using Stream replacementStream = replacement.Open();
            using StreamWriter writer = new(replacementStream, new UTF8Encoding(false));
            writer.Write(mastersDoc.Declaration + Environment.NewLine + mastersDoc.ToString(SaveOptions.DisableFormatting));
        }

        private static void RewriteEntry(string vsdxPath, string entryPath, string content) {
            using FileStream stream = File.Open(vsdxPath, FileMode.Open, FileAccess.ReadWrite);
            using ZipArchive archive = new(stream, ZipArchiveMode.Update);
            ZipArchiveEntry entry = archive.GetEntry(entryPath)!;
            entry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry(entryPath);
            using Stream replacementStream = replacement.Open();
            using StreamWriter writer = new(replacementStream, new UTF8Encoding(false));
            writer.Write(content.Replace("\r\n", "\n"));
        }

        private static string ReadEntry(string vsdxPath, string entryPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry(entryPath)!.Open();
            using StreamReader reader = new(stream, Encoding.UTF8);
            return reader.ReadToEnd();
        }

        private static string ReadConnectorGeometry(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement connectorShape = GetConnectorShape(pageDoc, ns);
            XElement geometry = connectorShape.Elements(ns + "Section").First(section => (string?)section.Attribute("N") == "Geometry");
            return geometry.ToString(SaveOptions.DisableFormatting);
        }

        private static string ReadConnectorLayoutFragments(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement connectorShape = GetConnectorShape(pageDoc, ns);

            return string.Concat(
                connectorShape.Elements(ns + "Cell")
                    .Where(cell => {
                        string? name = (string?)cell.Attribute("N");
                        return name == "TxtPinX" || name == "TxtPinY";
                    })
                    .Select(cell => cell.ToString(SaveOptions.DisableFormatting)),
                connectorShape.Elements(ns + "Section")
                    .Where(section => (string?)section.Attribute("N") == "Control")
                    .Select(section => section.ToString(SaveOptions.DisableFormatting)));
        }

        private static string ReadFirstShapeGeometry(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement shape = GetFirstShape(pageDoc, ns);
            XElement geometry = shape.Elements(ns + "Section").First(section => (string?)section.Attribute("N") == "Geometry");
            return geometry.ToString(SaveOptions.DisableFormatting);
        }

        private static string ReadFirstShapeLayoutFragments(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement shape = GetFirstShape(pageDoc, ns);

            return string.Concat(
                shape.Elements(ns + "Cell")
                    .Where(cell => {
                        string? name = (string?)cell.Attribute("N");
                        return name == "TxtPinX" || name == "TxtPinY";
                    })
                    .Select(cell => cell.ToString(SaveOptions.DisableFormatting)),
                shape.Elements(ns + "Section")
                    .Where(section => (string?)section.Attribute("N") == "User")
                    .Select(section => section.ToString(SaveOptions.DisableFormatting)));
        }

        private static XElement ReadConnectorTextElement(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement connectorShape = GetConnectorShape(pageDoc, ns);
            return new XElement(connectorShape.Element(ns + "Text")!);
        }

        private static XElement ReadFirstShapeTextElement(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement shape = GetFirstShape(pageDoc, ns);
            return new XElement(shape.Element(ns + "Text")!);
        }

        private static XElement ReadFirstShapeDataRow(string vsdxPath, string rowName) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement shape = GetFirstShape(pageDoc, ns);
            XElement propSection = shape.Elements(ns + "Section").Single(section => (string?)section.Attribute("N") == "Prop");
            return new XElement(propSection.Elements(ns + "Row").Single(row => (string?)row.Attribute("N") == rowName));
        }

        private static string ReadFirstPageSheetFragments(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/pages.xml")!.Open();
            XDocument pagesDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement pageSheet = GetFirstPageSheet(pagesDoc, ns);

            return string.Concat(
                pageSheet.Elements(ns + "Cell")
                    .Where(cell => {
                        string? name = (string?)cell.Attribute("N");
                        return name == "XRulerDensity" || name == "YRulerDensity";
                    })
                    .Select(cell => cell.ToString(SaveOptions.DisableFormatting)),
                pageSheet.Elements(ns + "Section")
                    .Where(section => (string?)section.Attribute("N") == "User")
                    .Select(section => section.ToString(SaveOptions.DisableFormatting)));
        }

        private static string ReadDocumentRootFragments(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/document.xml")!.Open();
            XDocument documentDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement root = documentDoc.Root!;

            return string.Concat(
                root.Attributes()
                    .Where(attribute => !attribute.IsNamespaceDeclaration &&
                                        attribute.Name.NamespaceName != "http://www.w3.org/XML/1998/namespace")
                    .Select(attribute => attribute.ToString()),
                root.Elements()
                    .Where(element =>
                        element.Name != ns + "DocumentSettings" &&
                        element.Name != ns + "Colors" &&
                        element.Name != ns + "FaceNames" &&
                        element.Name != ns + "StyleSheets")
                    .Select(element => element.ToString(SaveOptions.DisableFormatting)));
        }

        private static string ReadDocumentSettingsFragments(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/document.xml")!.Open();
            XDocument documentDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement settings = documentDoc.Root!.Element(ns + "DocumentSettings")!;

            return string.Concat(
                settings.Attributes()
                    .Where(attribute => attribute.Name.LocalName == "CustomSettingsFlag")
                    .Select(attribute => attribute.ToString()),
                settings.Elements()
                    .Where(element =>
                        element.Name == ns + "RelayoutAndRerouteUponOpen" ||
                        element.Name == ns + "CustomSettingsElement")
                    .Select(element => element.ToString(SaveOptions.DisableFormatting)));
        }

        private static string ReadStyleSheetsFragments(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/document.xml")!.Open();
            XDocument documentDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement styleSheets = documentDoc.Root!.Element(ns + "StyleSheets")!;

            return string.Concat(
                styleSheets.Attributes()
                    .Where(attribute => attribute.Name.LocalName == "CustomStyleRoot")
                    .Select(attribute => attribute.ToString()),
                styleSheets.Elements()
                    .Where(element => element.Name != ns + "StyleSheet")
                    .Select(element => element.ToString(SaveOptions.DisableFormatting)),
                styleSheets.Elements(ns + "StyleSheet")
                    .Where(styleSheet => (string?)styleSheet.Attribute("ID") is "0" or "1" or "2")
                    .Select(styleSheet => string.Concat(
                        styleSheet.Attributes()
                            .Where(attribute => attribute.Name.LocalName is "CustomAttr0" or "CustomAttr1" or "CustomAttr2")
                            .Select(attribute => attribute.ToString()),
                        styleSheet.Elements()
                            .Where(element =>
                                element.Name == ns + "Section" ||
                                (element.Name == ns + "Cell" && ((string?)element.Attribute("N") is "LineCap" or "LineColor")))
                            .Select(element => element.ToString(SaveOptions.DisableFormatting)))),
                styleSheets.Elements(ns + "StyleSheet")
                    .Where(styleSheet => (string?)styleSheet.Attribute("ID") == "7")
                    .Select(styleSheet => styleSheet.ToString(SaveOptions.DisableFormatting)));
        }

        private static string ReadColorsAndFaceNamesFragments(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/document.xml")!.Open();
            XDocument documentDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement root = documentDoc.Root!;
            XElement colors = root.Element(ns + "Colors")!;
            XElement faceNames = root.Element(ns + "FaceNames")!;

            return string.Concat(
                colors.Attributes()
                    .Where(attribute => attribute.Name.LocalName == "CustomPalette")
                    .Select(attribute => attribute.ToString()),
                colors.Elements()
                    .Select(element => element.ToString(SaveOptions.DisableFormatting)),
                faceNames.Attributes()
                    .Where(attribute => attribute.Name.LocalName == "CustomFonts")
                    .Select(attribute => attribute.ToString()),
                faceNames.Elements()
                    .Select(element => element.ToString(SaveOptions.DisableFormatting)));
        }

        private static string ReadFirstPageContentFragments(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement root = pageDoc.Root!;

            return string.Concat(
                root.Attributes()
                    .Where(attribute => attribute.Name.LocalName == "CustomPageContent")
                    .Select(attribute => attribute.ToString()),
                root.Elements()
                    .Where(element => element.Name != ns + "Shapes" && element.Name != ns + "Connects")
                    .Select(element => element.ToString(SaveOptions.DisableFormatting)));
        }

        private static string ReadFirstPageContainerFragments(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement root = pageDoc.Root!;
            XElement shapes = root.Element(ns + "Shapes")!;
            XElement connects = root.Element(ns + "Connects")!;

            return string.Concat(
                shapes.Attributes()
                    .Where(attribute => attribute.Name.LocalName == "ContainerFlag")
                    .Select(attribute => attribute.ToString()),
                shapes.Elements()
                    .Where(element => element.Name != ns + "Shape")
                    .Select(element => element.ToString(SaveOptions.DisableFormatting)),
                connects.Attributes()
                    .Where(attribute => attribute.Name.LocalName == "RoutingMode")
                    .Select(attribute => attribute.ToString()),
                connects.Elements()
                    .Where(element => element.Name != ns + "Connect")
                    .Select(element => element.ToString(SaveOptions.DisableFormatting)));
        }

        private static string ReadFirstShapeChildOrder(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement shape = GetFirstShape(pageDoc, ns);

            return string.Join("|",
                shape.Elements()
                    .Select(element => element.Name.LocalName switch {
                        "Cell" => $"Cell:{(string?)element.Attribute("N")}",
                        "Section" => $"Section:{(string?)element.Attribute("N")}",
                        _ => $"{element.Name.LocalName}:{(string?)element.Attribute("KeepOrder") ?? string.Empty}"
                    }));
        }

        private static string ReadShapesChildOrder(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement shapes = pageDoc.Root!.Element(ns + "Shapes")!;

            return string.Join("|",
                shapes.Elements()
                    .Select(element => string.Equals(element.Name.LocalName, "Shape", StringComparison.OrdinalIgnoreCase)
                        ? $"Shape:{(string?)element.Attribute("NameU")}"
                        : $"{element.Name.LocalName}:{(string?)element.Attribute("KeepOrder") ?? string.Empty}"));
        }

        private static string ReadConnectorShapeChildOrder(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement connectorShape = GetConnectorShape(pageDoc, ns);

            return string.Join("|",
                connectorShape.Elements()
                    .Select(element => element.Name.LocalName switch {
                        "Cell" => $"Cell:{(string?)element.Attribute("N")}",
                        "Section" => $"Section:{(string?)element.Attribute("N")}",
                        _ => $"{element.Name.LocalName}:{(string?)element.Attribute("KeepOrder") ?? string.Empty}"
                    }));
        }

        private static string ReadAdditionalConnectRows(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement connects = pageDoc.Root!.Element(ns + "Connects")!;

            return string.Concat(
                connects.Elements(ns + "Connect")
                    .Where(connect => (string?)connect.Attribute("FromCell") == "PinX")
                    .Select(connect => connect.ToString(SaveOptions.DisableFormatting)));
        }

        private static string ReadConnectRowOrder(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement connects = pageDoc.Root!.Element(ns + "Connects")!;

            return string.Join("|",
                connects.Elements(ns + "Connect")
                    .Select(connect => $"{(string?)connect.Attribute("FromCell")}:{(string?)connect.Attribute("CustomConnectFlag") ?? string.Empty}"));
        }

        private static string ReadConnectChildOrder(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement connects = pageDoc.Root!.Element(ns + "Connects")!;

            return string.Join("|",
                connects.Elements()
                    .Select(element => string.Equals(element.Name.LocalName, "Connect", StringComparison.OrdinalIgnoreCase)
                        ? $"Connect:{(string?)element.Attribute("FromCell")}"
                        : $"{element.Name.LocalName}:{(string?)element.Attribute("KeepOrder") ?? string.Empty}"));
        }

        private static string ReadNamespacedConnectAttributes(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement beginConnect = pageDoc.Root!
                .Element(ns + "Connects")!
                .Elements(ns + "Connect")
                .First(connect => (string?)connect.Attribute("FromCell") == "BeginX");

            return string.Concat(
                beginConnect.Attributes()
                    .Where(attribute => attribute.Name.NamespaceName == "urn:officeimo:connect-meta")
                    .Select(attribute => $"{{{attribute.Name.NamespaceName}}}{attribute.Name.LocalName}={attribute.Value}"));
        }

        private static string ReadConnectAttributeOrder(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement beginConnect = pageDoc.Root!
                .Element(ns + "Connects")!
                .Elements(ns + "Connect")
                .First(connect => (string?)connect.Attribute("FromCell") == "BeginX");

            return string.Join("|",
                beginConnect.Attributes()
                    .Where(attribute => !attribute.IsNamespaceDeclaration)
                    .Select(attribute => $"{{{attribute.Name.NamespaceName}}}{attribute.Name.LocalName}"));
        }

        private static string ReadFirstPageAttributes(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/pages.xml")!.Open();
            XDocument pagesDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement page = pagesDoc.Root!.Elements(ns + "Page").First();

            return string.Concat(
                page.Attributes()
                    .Where(attribute => attribute.Name.LocalName == "ReviewerID" || attribute.Name.LocalName == "AssociatedPage")
                    .Select(attribute => attribute.ToString()));
        }

        private static string ReadFirstMasterAttributes(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/masters/masters.xml")!.Open();
            XDocument mastersDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement master = mastersDoc.Root!.Elements(ns + "Master").First();

            return string.Concat(
                master.Attributes()
                    .Where(attribute => attribute.Name.LocalName == "PageType" || attribute.Name.LocalName == "IconFlags")
                    .Select(attribute => attribute.ToString()));
        }

        private static string ReadFirstMasterPageSheetFragments(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/masters/masters.xml")!.Open();
            XDocument mastersDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement pageSheet = mastersDoc.Root!.Elements(ns + "Master").First().Element(ns + "PageSheet")!;

            return string.Concat(
                pageSheet.Attributes()
                    .Where(attribute => attribute.Name.LocalName == "UniqueID")
                    .Select(attribute => attribute.ToString()),
                pageSheet.Elements(ns + "Cell")
                    .Where(cell => (string?)cell.Attribute("N") == "XRulerDensity")
                    .Select(cell => cell.ToString(SaveOptions.DisableFormatting)),
                pageSheet.Elements(ns + "Section")
                    .Where(section => (string?)section.Attribute("N") == "User")
                    .Select(section => section.ToString(SaveOptions.DisableFormatting)));
        }

        private static string ReadFirstMasterChildElements(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/masters/masters.xml")!.Open();
            XDocument mastersDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement master = mastersDoc.Root!.Elements(ns + "Master").First();

            return string.Concat(
                master.Elements()
                    .Where(element => element.Name != ns + "PageSheet" && element.Name != ns + "Rel")
                    .Select(element => element.ToString(SaveOptions.DisableFormatting)));
        }

        private static string ReadMastersRootFragments(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/masters/masters.xml")!.Open();
            XDocument mastersDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement root = mastersDoc.Root!;

            return string.Concat(
                root.Attributes()
                    .Where(attribute => !attribute.IsNamespaceDeclaration &&
                                        attribute.Name.NamespaceName != "http://www.w3.org/XML/1998/namespace")
                    .Select(attribute => attribute.ToString()),
                root.Elements()
                    .Where(element => element.Name != ns + "Master")
                    .Select(element => element.ToString(SaveOptions.DisableFormatting)));
        }

        private static string ReadFirstMasterContentFragments(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/masters/master1.xml")!.Open();
            XDocument masterDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement root = masterDoc.Root!;
            XElement shapes = root.Element(ns + "Shapes")!;

            return string.Concat(
                shapes.Elements(ns + "Shape").Skip(1).Select(shape => shape.ToString(SaveOptions.DisableFormatting)),
                root.Elements().Where(element => element.Name != ns + "Shapes").Select(element => element.ToString(SaveOptions.DisableFormatting)));
        }

        private static string ReadFirstMasterContentAttributes(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/masters/master1.xml")!.Open();
            XDocument masterDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement root = masterDoc.Root!;
            XElement shapes = root.Element(ns + "Shapes")!;

            return string.Concat(
                root.Attributes()
                    .Where(attribute => !attribute.IsNamespaceDeclaration &&
                                        attribute.Name.NamespaceName != "http://www.w3.org/XML/1998/namespace")
                    .Select(attribute => attribute.ToString()),
                shapes.Attributes()
                    .Where(attribute => !attribute.IsNamespaceDeclaration)
                    .Select(attribute => attribute.ToString()));
        }

        private static XElement GetFirstPageSheet(XDocument pagesDoc, XNamespace ns) {
            return pagesDoc.Root!
                .Elements(ns + "Page")
                .First()
                .Element(ns + "PageSheet")!;
        }

        private static XDocument NormalizeTheme(string xml) {
            return XDocument.Parse(xml);
        }
    }
}
