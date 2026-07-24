using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Tests {
    public class VisioTextStyleTests {
        [Fact]
        public void VisioLoadBoundsUntrustedFontFamilyNamesBeforeRendering() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.AddPage("Fonts")
                .AddRectangle(2, 4, 2.5, 1.1, "Bounded font")
                .ApplyTextStyle(new VisioTextStyle { FontFamily = new string('F', 10_000), Size = 12 });
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape shape = Assert.Single(loaded.Pages[0].Shapes);

            Assert.Equal(256, shape.TextStyle!.FontFamily!.Length);
            string svg = loaded.Pages[0].ToSvg();
            Assert.True(svg.Length < 100_000, $"Expected bounded SVG output but got {svg.Length} characters.");
        }

        [Fact]
        public void VisioTextStyleWritesCharacterParagraphAndTextBlockCells() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Text");
            page.AddRectangle(2, 4, 2.5, 1.1, "Styled text")
                .ApplyTextStyle(new VisioTextStyle {
                    FontFamily = "Aptos",
                    Color = Color.FromRgb(0x33, 0x66, 0x99),
                    Size = 13.5,
                    Bold = true,
                    Italic = true,
                    Underline = false,
                    HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                    VerticalAlignment = VisioTextVerticalAlignment.Middle,
                    LeftMargin = 0.08,
                    RightMargin = 0.09,
                    TopMargin = 0.1,
                    BottomMargin = 0.11,
                    BackgroundColor = Color.LightYellow,
                    BackgroundTransparency = 15,
                    TextPinX = 1.25,
                    TextPinY = -0.25,
                    TextWidth = 2.4,
                    TextHeight = 0.45,
                    TextLocPinX = 1.2,
                    TextLocPinY = 0.225,
                    TextAngle = 0.25
                });

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement faceName = Assert.Single(ReadXml(filePath, "visio/document.xml")
                .Descendants(ns + "FaceName"), element => (string?)element.Attribute("Name") == "Aptos");
            Assert.Equal("0", faceName.Attribute("ID")?.Value);

            XElement shape = FindShape(ReadXml(filePath, "visio/pages/page1.xml"), ns, "Styled text");

            Assert.Equal("0.08", Cell(shape, ns, "LeftMargin").Attribute("V")?.Value);
            Assert.Equal("0.09", Cell(shape, ns, "RightMargin").Attribute("V")?.Value);
            Assert.Equal("0.1", Cell(shape, ns, "TopMargin").Attribute("V")?.Value);
            Assert.Equal("0.11", Cell(shape, ns, "BottomMargin").Attribute("V")?.Value);
            Assert.Equal("1", Cell(shape, ns, "VerticalAlign").Attribute("V")?.Value);
            Assert.Equal("#FFFFE0", Cell(shape, ns, "TextBkgnd").Attribute("V")?.Value);
            Assert.Equal("15", Cell(shape, ns, "TextBkgndTrans").Attribute("V")?.Value);
            Assert.Equal("1.25", Cell(shape, ns, "TxtPinX").Attribute("V")?.Value);
            Assert.Equal("-0.25", Cell(shape, ns, "TxtPinY").Attribute("V")?.Value);
            Assert.Equal("2.4", Cell(shape, ns, "TxtWidth").Attribute("V")?.Value);
            Assert.Equal("0.45", Cell(shape, ns, "TxtHeight").Attribute("V")?.Value);
            Assert.Equal("1.2", Cell(shape, ns, "TxtLocPinX").Attribute("V")?.Value);
            Assert.Equal("0.225", Cell(shape, ns, "TxtLocPinY").Attribute("V")?.Value);
            Assert.Equal("0.25", Cell(shape, ns, "TxtAngle").Attribute("V")?.Value);

            XElement charSection = SingleSection(shape, ns, "Character");
            XElement charRow = Assert.Single(charSection.Elements(ns + "Row"));
            Assert.Equal("0", Cell(charRow, ns, "Font").Attribute("V")?.Value);
            Assert.Equal("#336699", Cell(charRow, ns, "Color").Attribute("V")?.Value);
            Assert.Equal("0.1875", Cell(charRow, ns, "Size").Attribute("V")?.Value);
            Assert.Equal("PT", Cell(charRow, ns, "Size").Attribute("U")?.Value);
            Assert.Equal("3", Cell(charRow, ns, "Style").Attribute("V")?.Value);

            XElement paraSection = SingleSection(shape, ns, "Paragraph");
            XElement paraRow = Assert.Single(paraSection.Elements(ns + "Row"));
            Assert.Equal("1", Cell(paraRow, ns, "HorzAlign").Attribute("V")?.Value);
        }

        [Fact]
        public void VisioTextStyleRoundTripsWithoutDuplicatingModeledSections() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("RoundTrip");
            page.AddRectangle(2, 4, 2.5, 1.1, "Round tripped")
                .ApplyTextStyle(new VisioTextStyle {
                    FontFamily = "Consolas",
                    Color = Color.Crimson,
                    Size = 12,
                    Bold = true,
                    HorizontalAlignment = VisioTextHorizontalAlignment.Right,
                    VerticalAlignment = VisioTextVerticalAlignment.Bottom,
                    LeftMargin = 0.12,
                    BackgroundColor = Color.LightCyan,
                    BackgroundTransparency = 25,
                    TextPinX = 1.2,
                    TextPinY = -0.2,
                    TextWidth = 2.1,
                    TextHeight = 0.5,
                    TextLocPinX = 1.05,
                    TextLocPinY = 0.25,
                    TextAngle = 0.15
                });
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape loadedShape = Assert.Single(loaded.Pages[0].Shapes, shape => shape.Text == "Round tripped");

            Assert.NotNull(loadedShape.TextStyle);
            Assert.Equal("Consolas", loadedShape.TextStyle!.FontFamily);
            Assert.Equal(Color.Crimson, loadedShape.TextStyle.Color);
            Assert.Equal(12, loadedShape.TextStyle.Size);
            Assert.True(loadedShape.TextStyle.Bold);
            Assert.False(loadedShape.TextStyle.Italic);
            Assert.False(loadedShape.TextStyle.Underline);
            Assert.Equal(VisioTextHorizontalAlignment.Right, loadedShape.TextStyle.HorizontalAlignment);
            Assert.Equal(VisioTextVerticalAlignment.Bottom, loadedShape.TextStyle.VerticalAlignment);
            Assert.Equal(0.12, loadedShape.TextStyle.LeftMargin);
            Assert.Equal(Color.LightCyan, loadedShape.TextStyle.BackgroundColor);
            Assert.Equal(25, loadedShape.TextStyle.BackgroundTransparency);
            Assert.Equal(1.2, loadedShape.TextStyle.TextPinX);
            Assert.Equal(-0.2, loadedShape.TextStyle.TextPinY);
            Assert.Equal(2.1, loadedShape.TextStyle.TextWidth);
            Assert.Equal(0.5, loadedShape.TextStyle.TextHeight);
            Assert.Equal(1.05, loadedShape.TextStyle.TextLocPinX);
            Assert.Equal(0.25, loadedShape.TextStyle.TextLocPinY);
            Assert.Equal(0.15, loadedShape.TextStyle.TextAngle);

            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Empty(VisioValidator.Validate(savedPath));
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement savedShape = FindShape(ReadXml(savedPath, "visio/pages/page1.xml"), ns, "Round tripped");
            Assert.Single(savedShape.Elements(ns + "Section"), section => (string?)section.Attribute("N") == "Character");
            Assert.Single(savedShape.Elements(ns + "Section"), section => (string?)section.Attribute("N") == "Paragraph");
            Assert.Single(ReadXml(savedPath, "visio/document.xml")
                .Descendants(ns + "FaceName"), element => (string?)element.Attribute("Name") == "Consolas");
        }

        [Fact]
        public void ConnectorLabelTextStyleWritesCharacterParagraphAndTextBlockCells() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("ConnectorText");
            VisioShape source = page.AddRectangle(1, 4, 1.5, 0.7, "Source");
            VisioShape target = page.AddRectangle(5, 4, 1.5, 0.7, "Target");
            page.AddConnector(source, target, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(3.25, 4.45, width: 1.4, height: 0.35)
                .ApplyTextStyle(new VisioTextStyle {
                    FontFamily = "Aptos",
                    Color = Color.DodgerBlue,
                    Size = 9.5,
                    Bold = true,
                    Underline = true,
                    HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                    VerticalAlignment = VisioTextVerticalAlignment.Middle,
                    LeftMargin = 0.03,
                    RightMargin = 0.04,
                    TopMargin = 0.05,
                    BottomMargin = 0.06,
                    BackgroundColor = Color.White,
                    BackgroundTransparency = 0,
                    TextPinX = 9
                })
                .Label = "Approved";

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement connector = FindShape(ReadXml(filePath, "visio/pages/page1.xml"), ns, "Approved");

            Assert.Equal("0.03", Cell(connector, ns, "LeftMargin").Attribute("V")?.Value);
            Assert.Equal("0.04", Cell(connector, ns, "RightMargin").Attribute("V")?.Value);
            Assert.Equal("0.05", Cell(connector, ns, "TopMargin").Attribute("V")?.Value);
            Assert.Equal("0.06", Cell(connector, ns, "BottomMargin").Attribute("V")?.Value);
            Assert.Equal("1", Cell(connector, ns, "VerticalAlign").Attribute("V")?.Value);
            Assert.Equal("#FFFFFF", Cell(connector, ns, "TextBkgnd").Attribute("V")?.Value);
            Assert.Equal("0", Cell(connector, ns, "TextBkgndTrans").Attribute("V")?.Value);
            Assert.Equal("3.25", Cell(connector, ns, "TxtPinX").Attribute("V")?.Value);
            Assert.Equal("4.45", Cell(connector, ns, "TxtPinY").Attribute("V")?.Value);
            Assert.Equal("1.4", Cell(connector, ns, "TxtWidth").Attribute("V")?.Value);
            Assert.Equal("0.35", Cell(connector, ns, "TxtHeight").Attribute("V")?.Value);

            XElement charSection = SingleSection(connector, ns, "Character");
            XElement charRow = Assert.Single(charSection.Elements(ns + "Row"));
            Assert.Equal("0", Cell(charRow, ns, "Font").Attribute("V")?.Value);
            Assert.Equal("#1E90FF", Cell(charRow, ns, "Color").Attribute("V")?.Value);
            Assert.Equal("0.131944444444444", Cell(charRow, ns, "Size").Attribute("V")?.Value);
            Assert.Equal("PT", Cell(charRow, ns, "Size").Attribute("U")?.Value);
            Assert.Equal("5", Cell(charRow, ns, "Style").Attribute("V")?.Value);

            XElement paraSection = SingleSection(connector, ns, "Paragraph");
            XElement paraRow = Assert.Single(paraSection.Elements(ns + "Row"));
            Assert.Equal("1", Cell(paraRow, ns, "HorzAlign").Attribute("V")?.Value);
            Assert.Single(ReadXml(filePath, "visio/document.xml")
                .Descendants(ns + "FaceName"), element => (string?)element.Attribute("Name") == "Aptos");
        }

        [Fact]
        public void ConnectorLabelTextStyleRoundTripsWithoutDuplicatingModeledSections() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("ConnectorRoundTrip");
            VisioShape source = page.AddRectangle(1, 4, 1.5, 0.7, "Source");
            VisioShape target = page.AddRectangle(5, 4, 1.5, 0.7, "Target");
            page.AddConnector(source, target, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(3.25, 4.45, width: 1.4, height: 0.35)
                .ApplyTextStyle(new VisioTextStyle {
                    FontFamily = "Consolas",
                    Color = Color.Crimson,
                    Size = 10,
                    Italic = true,
                    HorizontalAlignment = VisioTextHorizontalAlignment.Right,
                    VerticalAlignment = VisioTextVerticalAlignment.Bottom,
                    LeftMargin = 0.07,
                    BackgroundColor = Color.LightCyan,
                    BackgroundTransparency = 35
                })
                .Label = "Denied";
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector loadedConnector = Assert.Single(loaded.Pages[0].Connectors);

            Assert.NotNull(loadedConnector.TextStyle);
            Assert.Equal("Consolas", loadedConnector.TextStyle!.FontFamily);
            Assert.Equal(Color.Crimson, loadedConnector.TextStyle.Color);
            Assert.Equal(10, loadedConnector.TextStyle.Size);
            Assert.False(loadedConnector.TextStyle.Bold);
            Assert.True(loadedConnector.TextStyle.Italic);
            Assert.False(loadedConnector.TextStyle.Underline);
            Assert.Equal(VisioTextHorizontalAlignment.Right, loadedConnector.TextStyle.HorizontalAlignment);
            Assert.Equal(VisioTextVerticalAlignment.Bottom, loadedConnector.TextStyle.VerticalAlignment);
            Assert.Equal(0.07, loadedConnector.TextStyle.LeftMargin);
            Assert.Equal(Color.LightCyan, loadedConnector.TextStyle.BackgroundColor);
            Assert.Equal(35, loadedConnector.TextStyle.BackgroundTransparency);
            Assert.Equal(3.25, loadedConnector.LabelPlacement!.AbsolutePinX);
            Assert.Equal(4.45, loadedConnector.LabelPlacement!.AbsolutePinY);

            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.Empty(VisioValidator.Validate(savedPath));
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement savedConnector = FindShape(ReadXml(savedPath, "visio/pages/page1.xml"), ns, "Denied");
            Assert.Single(savedConnector.Elements(ns + "Section"), section => (string?)section.Attribute("N") == "Character");
            Assert.Single(savedConnector.Elements(ns + "Section"), section => (string?)section.Attribute("N") == "Paragraph");
            Assert.Single(ReadXml(savedPath, "visio/document.xml")
                .Descendants(ns + "FaceName"), element => (string?)element.Attribute("Name") == "Consolas");
        }

        [Fact]
        public void NativeComplexConnectorCharSectionsArePreservedWithoutModeledDuplicates() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("NativeConnector");
            VisioShape source = page.AddRectangle(1, 4, 1.5, 0.7, "Source");
            VisioShape target = page.AddRectangle(5, 4, 1.5, 0.7, "Target");
            page.AddConnector(source, target, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left).Label = "Native connector text";
            document.Save();

            RewritePage(filePath, pageXml => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement connector = FindShape(pageXml, ns, "Native connector text");
                connector.Elements(ns + "Section").Where(section => (string?)section.Attribute("N") == "Char").Remove();
                connector.Add(new XElement(ns + "Section",
                    new XAttribute("N", "Char"),
                    new XElement(ns + "Row",
                        new XAttribute("IX", "0"),
                        new XElement(ns + "Cell", new XAttribute("N", "Color"), new XAttribute("V", "#112233")),
                        new XElement(ns + "Cell", new XAttribute("N", "Case"), new XAttribute("V", "1")))));
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement savedConnector = FindShape(ReadXml(savedPath, "visio/pages/page1.xml"), v, "Native connector text");
            XElement charSection = Assert.Single(savedConnector.Elements(v + "Section"), section => (string?)section.Attribute("N") == "Char");
            XElement row = Assert.Single(charSection.Elements(v + "Row"));
            Assert.Equal("#112233", Cell(row, v, "Color").Attribute("V")?.Value);
            Assert.Equal("1", Cell(row, v, "Case").Attribute("V")?.Value);
            Assert.Empty(VisioValidator.Validate(savedPath));
        }

        [Fact]
        public void NativeComplexCharSectionsArePreservedWithoutModeledDuplicates() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Native");
            page.AddRectangle(2, 4, 2.5, 1.1, "Native rich text");
            document.Save();

            RewritePage(filePath, pageXml => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement shape = FindShape(pageXml, ns, "Native rich text");
                shape.Elements(ns + "Section").Where(section => (string?)section.Attribute("N") == "Char").Remove();
                shape.Add(new XElement(ns + "Section",
                    new XAttribute("N", "Char"),
                    new XElement(ns + "Row",
                        new XAttribute("IX", "0"),
                        new XElement(ns + "Cell", new XAttribute("N", "Color"), new XAttribute("V", "#112233")),
                        new XElement(ns + "Cell", new XAttribute("N", "Case"), new XAttribute("V", "1")))));
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement savedShape = FindShape(ReadXml(savedPath, "visio/pages/page1.xml"), v, "Native rich text");
            XElement charSection = Assert.Single(savedShape.Elements(v + "Section"), section => (string?)section.Attribute("N") == "Char");
            XElement row = Assert.Single(charSection.Elements(v + "Row"));
            Assert.Equal("#112233", Cell(row, v, "Color").Attribute("V")?.Value);
            Assert.Equal("1", Cell(row, v, "Case").Attribute("V")?.Value);
            Assert.Empty(VisioValidator.Validate(savedPath));
        }

        [Fact]
        public void ApplyTextStyleSelectionUsesDetachedCopies() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Selection");
            VisioShape first = page.AddRectangle(2, 4, 2, 1, "One");
            VisioShape second = page.AddRectangle(5, 4, 2, 1, "Two");
            VisioTextStyle style = new() {
                Size = 11,
                Bold = true,
                HorizontalAlignment = VisioTextHorizontalAlignment.Justify
            };

            page.SelectShapes(_ => true).ApplyTextStyle(style);
            style.Size = 20;
            first.TextStyle!.Bold = false;

            Assert.NotSame(style, first.TextStyle);
            Assert.NotSame(first.TextStyle, second.TextStyle);
            Assert.Equal(11, first.TextStyle.Size);
            Assert.Equal(11, second.TextStyle!.Size);
            Assert.False(first.TextStyle.Bold);
            Assert.True(second.TextStyle.Bold);
            Assert.Equal(VisioTextHorizontalAlignment.Justify, second.TextStyle.HorizontalAlignment);
        }

        [Fact]
        public void ApplyTextStyleConnectorSelectionUsesDetachedCopies() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("ConnectorSelection");
            VisioShape source = page.AddRectangle(1, 4, 1.5, 0.7, "Source");
            VisioShape middle = page.AddRectangle(3, 4, 1.5, 0.7, "Middle");
            VisioShape target = page.AddRectangle(5, 4, 1.5, 0.7, "Target");
            VisioConnector first = page.AddConnector(source, middle, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            VisioConnector second = page.AddConnector(middle, target, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            VisioTextStyle style = new() {
                Size = 8,
                Bold = true,
                HorizontalAlignment = VisioTextHorizontalAlignment.Center
            };

            new VisioConnectorSelection(new[] { first, second }).ApplyTextStyle(style);
            style.Size = 20;
            first.TextStyle!.Bold = false;

            Assert.NotSame(style, first.TextStyle);
            Assert.NotSame(first.TextStyle, second.TextStyle);
            Assert.Equal(8, first.TextStyle.Size);
            Assert.Equal(8, second.TextStyle!.Size);
            Assert.False(first.TextStyle.Bold);
            Assert.True(second.TextStyle.Bold);
            Assert.Equal(VisioTextHorizontalAlignment.Center, second.TextStyle.HorizontalAlignment);
        }

        private static XDocument ReadXml(string vsdxPath, string entryName) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException("Missing " + entryName);
            using Stream stream = entry.Open();
            return XDocument.Load(stream);
        }

        private static void RewritePage(string vsdxPath, Action<XDocument> transform) {
            using FileStream stream = File.Open(vsdxPath, FileMode.Open, FileAccess.ReadWrite);
            using ZipArchive archive = new(stream, ZipArchiveMode.Update);
            ZipArchiveEntry pageEntry = archive.GetEntry("visio/pages/page1.xml") ?? throw new InvalidOperationException("Missing page1.xml");
            XDocument pageXml;
            using (Stream pageStream = pageEntry.Open()) {
                pageXml = XDocument.Load(pageStream);
            }

            transform(pageXml);
            pageEntry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry("visio/pages/page1.xml");
            using Stream replacementStream = replacement.Open();
            pageXml.Save(replacementStream, SaveOptions.DisableFormatting);
        }

        private static XElement FindShape(XDocument pageXml, XNamespace ns, string text) {
            return pageXml.Descendants(ns + "Shape")
                .Single(shape => string.Equals(shape.Element(ns + "Text")?.Value, text, StringComparison.Ordinal));
        }

        private static XElement SingleSection(XElement shape, XNamespace ns, string name) {
            return shape.Elements(ns + "Section")
                .Single(section => string.Equals(section.Attribute("N")?.Value, name, StringComparison.Ordinal));
        }

        private static XElement Cell(XElement parent, XNamespace ns, string name) {
            return parent.Elements(ns + "Cell")
                .Single(cell => string.Equals(cell.Attribute("N")?.Value, name, StringComparison.Ordinal));
        }
    }
}
