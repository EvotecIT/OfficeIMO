using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioArcMasterScalingTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";
    private const double OriginX = 2, OriginY = 6, Density = 40;

    // Compare the public rendered curve with the affine transform of the original,
    // rather than duplicating the native axis decomposition in this test.
    [Theory]
    [InlineData("ArcTo", 0, 0, 3, 0, .8, 0, 0, 1)]
    [InlineData("ArcTo", 0, 0, 3, 0, -.8, 0, 0, 1)]
    [InlineData("ArcTo", 3, 0, 0, 0, .8, 0, 0, 1)]
    [InlineData("ArcTo", 0, 0, 3, 4, 1, 0, 0, 1)]
    [InlineData("ArcTo", 0, 0, 3, 0, 2, 0, 0, 1)]
    [InlineData("EllipticalArcTo", .7901955449320697, -1.0561718779622504, 3.0800942105485363, 1.4020841117064462, 2.305372979190097, -.10529795926452512, .7853981633974483, 2)]
    [InlineData("RelEllipticalArcTo", .7901955449320697, -1.0561718779622504, 3.0800942105485363, 1.4020841117064462, 2.305372979190097, -.10529795926452512, .7853981633974483, 2)]
    [InlineData("EllipticalArcTo", .9404585238123429, .04663478642817387, 1.7775133875500637, -.08768440087340734, 1.2390089903610162, -.46339549465045293, -.3, .5)]
    public void NonuniformMasterCreationRetainsEditableAffineArcsThroughCopiesAndReopening(
        string type, double startX, double startY, double endX, double endY, double a, double b, double angle, double ratio) {
        foreach (var scale in new[] { (X: 2D, Y: .5D), (X: .5D, Y: 2D) }) {
            VisioDocument document = LoadMaster(type, startX, startY, endX, endY, a, b, angle, ratio);
            VisioMaster master = document.GetMaster("Arc");
            var originalPage = document.AddPage("Original", 12, 12);
            originalPage.AddShape("original", master, OriginX, OriginY, 4, 4);
            XElement originalMaster = SavedMaster(document);
            List<(double X, double Y)> originalPoints = Points(originalPage);
            var page = document.AddPage("Scaled", 12, 12);
            // Both overloads must use the same geometry owner.
            if (scale.X > scale.Y) page.AddShape("scaled", master, OriginX, OriginY, 4 * scale.X, 4 * scale.Y);
            else page.AddShape("scaled", "Arc", OriginX, OriginY, 4 * scale.X, 4 * scale.Y);
            document.DuplicatePage(page, "Copy");

            IEnumerable<VisioDocument> reopenedDocuments = new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) };
            // Relative geometry has no Visio 2003 XML row mapping. Its editable
            // native contract is Open XML, while absolute arcs support both formats.
            if (!type.StartsWith("Rel", StringComparison.Ordinal)) reopenedDocuments = reopenedDocuments.Append(LoadXml(document.ToLegacyXmlResult().Value));
            foreach (VisioDocument reopened in reopenedDocuments) {
                Assert.True(XNode.DeepEquals(originalMaster, SavedMaster(reopened)), "Resizing must leave the source master and its formulas intact.");
                byte[] before = reopened.ToLegacyXmlResult().Value;
                foreach (string name in new[] { "Scaled", "Copy" }) {
                    VisioPage scaledPage = reopened.Pages.Single(p => p.Name == name);
                    List<(double X, double Y)> points = Points(scaledPage);
                    Assert.Equal(originalPoints.Count, points.Count);
                    for (int i = 0; i < points.Count; i++) {
                        // Public SVG coordinates are rounded to .001 pixels. Allow
                        // the source and scaled output's combined rounding at 40 ppi.
                        double expectedX = OriginX + (originalPoints[i].X - OriginX) * scale.X;
                        double expectedY = OriginY + (originalPoints[i].Y - OriginY) * scale.Y;
                        Assert.InRange(points[i].X, expectedX - .00005, expectedX + .00005);
                        Assert.InRange(points[i].Y, expectedY - .00005, expectedY + .00005);
                    }
                    XElement geometry = ReadPart(reopened, "visio/pages/page" + (reopened.Pages.ToList().IndexOf(scaledPage) + 1) + ".xml").Descendants(Modern + "Section").Single(s => (string?)s.Attribute("N") == "Geometry");
                    XElement arc = geometry.Elements(Modern + "Row").Single(row => (string?)row.Attribute("IX") == "9");
                    Assert.Equal(type.StartsWith("Rel", StringComparison.Ordinal) ? "RelEllipticalArcTo" : "EllipticalArcTo", (string?)arc.Attribute("T"));
                    Assert.InRange(double.Parse((string)arc.Elements(Modern + "Cell").Single(c => (string?)c.Attribute("N") == "D").Attribute("V")!, CultureInfo.InvariantCulture), 1, 1000);
                    Assert.Contains(arc.Elements(Modern + "Cell"), c => (string?)c.Attribute("N") == "C");
                    Assert.Equal("1", (string?)geometry.Elements(Modern + "Cell").Single(c => (string?)c.Attribute("N") == "NoFill").Attribute("V"));
                    if (type == "ArcTo") Assert.Null(arc.Elements(Modern + "Cell").Single(c => (string?)c.Attribute("N") == "A").Attribute("F"));
                }
                Assert.Equal(before, reopened.ToLegacyXmlResult().Value);
            }
        }
    }

    [Fact]
    public void UniformArcsAndZeroBowRetainTheirNativeRowsAndApplicableFormulas() {
        foreach (double bow in new[] { 0D, .8D }) {
            VisioDocument document = LoadMaster("ArcTo", 0, 0, 3, 0, bow, 0, 0, 1);
            document.AddPage("Scaled").AddShape("scaled", "Arc", OriginX, OriginY, 8, bow == 0 ? 2 : 8);
            XElement arc = SavedGeometry(document, "Scaled").Element(Legacy + "ArcTo")!;
            Assert.NotNull(arc);
            Assert.Equal(bow * 2, (double)arc.Element(Legacy + "A")!, 8);
            Assert.Equal("Width*" + (bow / 4).ToString("R", CultureInfo.InvariantCulture), (string?)arc.Element(Legacy + "A")!.Attribute("F"));
        }
    }

    [Fact]
    public void ArcScalingReadsFormulaOnlyCoordinatesAndIgnoresDeletedRows() {
        VisioDocument document = LoadMaster("ArcTo", 0, 0, 3, 0, .8, 0, 0, 1);
        document.AddPage("Original").AddShape("original", "Arc", OriginX, OriginY, 4, 4);
        XElement xml = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value)).Root!;
        XElement geometry = xml.Descendants(Legacy + "Master").Descendants(Legacy + "Geom").Single();
        XElement move = geometry.Element(Legacy + "MoveTo")!;
        move.Element(Legacy + "X")!.ReplaceWith(new XElement(Legacy + "X", new XAttribute("F", "Width*0.25")));
        move.AddAfterSelf(new XElement(Legacy + "LineTo", new XAttribute("IX", "6"), new XAttribute("Del", "1"), Cell("X", 99), Cell("Y", 99)),
            new XElement(Legacy + "ArcTo", new XAttribute("IX", "7"), new XAttribute("Del", "1"), Cell("X", 88), Cell("Y", 88), Cell("A", 20)));
        document = LoadXml(Encoding.UTF8.GetBytes(xml.ToString()));
        document.AddPage("Scaled", 12, 12).AddShape("scaled", "Arc", OriginX, OriginY, 8, 2);
        XElement scaled = SavedGeometry(document, "Scaled");
        Assert.Equal(2, (double)scaled.Element(Legacy + "MoveTo")!.Element(Legacy + "X")!);
        XElement arc = scaled.Elements().Single(row => (string?)row.Attribute("IX") == "9");
        Assert.Equal(4, (double)arc.Element(Legacy + "A")!, 8);
        Assert.Equal(-.4, (double)arc.Element(Legacy + "B")!, 8);
        Assert.Equal("ArcTo", scaled.Elements().Single(row => (string?)row.Attribute("IX") == "7").Name.LocalName);
        Assert.Equal(20, (double)scaled.Elements().Single(row => (string?)row.Attribute("IX") == "7").Element(Legacy + "A")!);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1001)]
    public void InvalidEllipticalProfilesFailBeforeInsertionAndLeaveMasterUnchanged(double ratio) {
        VisioDocument document = LoadMaster("EllipticalArcTo", 0, 0, 3, 0, 1.5, -.8, .3, ratio);
        document.AddPage("Original").AddShape("original", "Arc", OriginX, OriginY, 4, 4);
        VisioPage page = document.AddPage("Scaled");
        XElement before = SavedMaster(document);
        Assert.Throws<NotSupportedException>(() => page.AddShape("failed", "Arc", OriginX, OriginY, 8, 2));
        Assert.Empty(page.Shapes);
        Assert.True(XNode.DeepEquals(before, SavedMaster(document)));
    }

    [Fact]
    public void ExcessiveAspectRatioFailsBeforeInsertionWithoutChangingTheSource() {
        VisioDocument document = LoadMaster("ArcTo", 0, 0, 3, 0, .8, 0, 0, 1);
        document.AddPage("Original").AddShape("original", "Arc", OriginX, OriginY, 4, 4);
        VisioPage page = document.AddPage("Scaled");
        XElement before = SavedMaster(document);
        Assert.Throws<NotSupportedException>(() => page.AddShape("failed", "Arc", OriginX, OriginY, 4004, 4));
        Assert.Empty(page.Shapes);
        Assert.True(XNode.DeepEquals(before, SavedMaster(document)));
    }

    [Fact]
    public void RepurposedBowCellDoesNotRetainProducerErrorWhenItsCacheCoincides() {
        VisioDocument document = LoadMaster("ArcTo", 0, 0, 1.6, 0, .8, 0, 0, 1);
        document.AddPage("Original").AddShape("original", "Arc", OriginX, OriginY, 4, 4);
        XElement xml = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value)).Root!;
        XElement bow = xml.Descendants(Legacy + "Master").Descendants(Legacy + "A").Single();
        bow.Attribute("F")!.Remove();
        bow.SetAttributeValue("Err", "producer-error");
        document = LoadXml(Encoding.UTF8.GetBytes(xml.ToString()));
        document.AddPage("Scaled").AddShape("scaled", "Arc", OriginX, OriginY, 4, 8);
        foreach (VisioDocument reopened in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) }) {
            XElement control = SavedGeometry(reopened, "Scaled").Element(Legacy + "EllipticalArcTo")!.Element(Legacy + "A")!;
            Assert.Equal(.8, (double)control, 8);
            Assert.Null(control.Attribute("Err"));
            XElement masterBow = XDocument.Load(new MemoryStream(reopened.ToLegacyXmlResult().Value)).Descendants(Legacy + "Master").Descendants(Legacy + "ArcTo").Single().Element(Legacy + "A")!;
            Assert.Equal("producer-error", (string?)masterBow.Attribute("Err"));
        }
    }

    [Theory]
    [InlineData(8, 2)]
    [InlineData(4, 8)]
    public void GeometryFormulasAreComparedAgainstFinalInstancePlacement(double width, double height) {
        VisioDocument document = LoadMaster("ArcTo", 2, 0, 3, 0, .8, 0, 0, 1);
        document.AddPage("Original").AddShape("original", "Arc", 1, 6, 4, 4);
        XElement xml = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value)).Root!;
        XElement masterShape = xml.Descendants(Legacy + "Master").Descendants(Legacy + "Shape").Single();
        masterShape.Element(Legacy + "XForm")!.Element(Legacy + "PinX")!.Value = "1";
        masterShape.Descendants(Legacy + "MoveTo").Single().Element(Legacy + "X")!.SetAttributeValue("F", "PinX*2");
        document = LoadXml(Encoding.UTF8.GetBytes(xml.ToString()));
        document.AddPage("Scaled").AddShape("scaled", "Arc", 7, 6, width, height);
        XElement coordinate = SavedGeometry(document, "Scaled").Element(Legacy + "MoveTo")!.Element(Legacy + "X")!;
        Assert.Equal(width / 2, (double)coordinate, 8);
        Assert.Null(coordinate.Attribute("F"));
        Assert.Equal("PinX*2", (string?)XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value)).Descendants(Legacy + "Master").Descendants(Legacy + "MoveTo").Single().Element(Legacy + "X")!.Attribute("F"));
    }

    [Theory]
    [InlineData(8, 2)]
    [InlineData(8, 8)]
    public void RelativeCoordinatesKeepTheirFractionsWithoutRetainingStaleInstanceFormulas(double width, double height) {
        VisioDocument document = LoadMaster("RelEllipticalArcTo", 1, 0, 3, 0, 1.5, -.8, .3, 2);
        document = RewriteMaster(document, native => {
            XElement shape = native.Descendants(Modern + "Shape").Single();
            shape.Elements(Modern + "Cell").Single(c => (string?)c.Attribute("N") == "PinX").SetAttributeValue("V", "1");
            XElement geometry = shape.Elements(Modern + "Section").Single(s => (string?)s.Attribute("N") == "Geometry");
            XElement move = geometry.Elements(Modern + "Row").First();
            move.Elements(Modern + "Cell").Single(c => (string?)c.Attribute("N") == "X").SetAttributeValue("F", "PinX*0.25");
            move.AddAfterSelf(new XElement(Modern + "Row", new XAttribute("T", "RelLineTo"), new XAttribute("IX", "6"),
                NativeCell("X", .25, "Width/Height*0.25"), NativeCell("Y", .25, "Height*0.0625")));
            XElement arc = geometry.Elements(Modern + "Row").Single(r => (string?)r.Attribute("IX") == "9");
            foreach (var cell in arc.Elements(Modern + "Cell")) {
                string? name = (string?)cell.Attribute("N");
                if (name == "X" || name == "A") cell.SetAttributeValue("F", "Width*" + ((double)cell.Attribute("V")! / 4).ToString("R", CultureInfo.InvariantCulture));
                if (name == "B") cell.SetAttributeValue("F", "Height*-0.05");
                if (name == "Y") cell.SetAttributeValue("F", "Height*0");
            }
        });
        document.AddPage("Original").AddShape("original", "Arc", 1, 6, 4, 4);
        XElement originalMaster = SavedMaster(document);
        document.AddPage("Scaled").AddShape("scaled", "Arc", 7, 6, width, height);
        foreach (VisioDocument reopened in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) }) {
            Assert.True(XNode.DeepEquals(originalMaster, SavedMaster(reopened)));
            XElement geometry = ReadPart(reopened, "visio/pages/page3.xml").Descendants(Modern + "Section").Single(s => (string?)s.Attribute("N") == "Geometry");
            XElement move = geometry.Elements(Modern + "Row").Single(r => (string?)r.Attribute("IX") == "5");
            XElement line = geometry.Elements(Modern + "Row").Single(r => (string?)r.Attribute("IX") == "6");
            XElement arc = geometry.Elements(Modern + "Row").Single(r => (string?)r.Attribute("IX") == "9");
            AssertCell(move, "X", .25, null);
            AssertCell(line, "X", .25, width == height ? "Width/Height*0.25" : null);
            AssertCell(line, "Y", .25, null);
            AssertCell(arc, "X", .75, null);
            AssertCell(arc, "A", .375, null);
            AssertCell(arc, "B", -.2, null);
            AssertCell(arc, "Y", 0, "Height*0");
        }

        static XElement NativeCell(string name, double value, string formula) => new(Modern + "Cell", new XAttribute("N", name), new XAttribute("V", value.ToString("R", CultureInfo.InvariantCulture)), new XAttribute("F", formula));
        static void AssertCell(XElement row, string name, double value, string? formula) {
            XElement cell = row.Elements(Modern + "Cell").Single(c => (string?)c.Attribute("N") == name);
            Assert.Equal(value, (double)cell.Attribute("V")!, 8);
            Assert.Equal(formula, (string?)cell.Attribute("F"));
        }
    }

    private static List<(double X, double Y)> Points(VisioPage page) {
        XDocument xml = XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = Density, RenderStencilArtwork = false }));
        XElement path = Assert.Single(xml.Descendants(Svg + "path"), p => (string?)p.Attribute("data-officeimo-preserved-geometry") == "true");
        string[] parts = path.Attribute("d")!.Value.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
        var result = new List<(double X, double Y)>();
        for (int i = 0; i < parts.Length; i++) {
            if (parts[i] != "M" && parts[i] != "L") continue;
            double x = double.Parse(parts[++i], CultureInfo.InvariantCulture) / Density;
            double y = 12 - double.Parse(parts[++i], CultureInfo.InvariantCulture) / Density;
            result.Add((x, y));
        }
        return result;
    }

    private static XElement SavedMaster(VisioDocument document) => ReadPart(document, "visio/masters/master1.xml").Root!;
    private static XElement SavedGeometry(VisioDocument document, string page) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value)).Descendants(Legacy + "Page").Single(p => (string?)p.Attribute("Name") == page).Descendants(Legacy + "Geom").Single();
    private static VisioDocument LoadXml(byte[] xml) => VisioDocument.LoadLegacyXml(new MemoryStream(xml)).Value;

    private static XDocument ReadPart(VisioDocument document, string name) {
        using var zip = new ZipArchive(new MemoryStream(document.ToBytes()), ZipArchiveMode.Read);
        using var part = zip.GetEntry(name)!.Open();
        return XDocument.Load(part);
    }

    private static VisioDocument LoadMaster(string type, double startX, double startY, double endX, double endY, double a, double b, double angle, double ratio) {
        bool relative = type.StartsWith("Rel", StringComparison.Ordinal);
        XElement arc = new(Legacy + (relative ? "EllipticalArcTo" : type), new XAttribute("IX", "9"), Cell("X", endX), Cell("Y", endY), Cell("A", a));
        if (type == "ArcTo") arc.Element(Legacy + "A")!.SetAttributeValue("F", "Width*" + (a / 4).ToString("R", CultureInfo.InvariantCulture));
        else arc.Add(Cell("B", b), Cell("C", angle), Cell("D", ratio));
        XElement shape = new(Legacy + "Shape", new XAttribute("ID", "1"), new XAttribute("Type", "Shape"),
            new XElement(Legacy + "XForm", Cell("PinX", 0), Cell("PinY", 0), Cell("Width", 4), Cell("Height", 4), Cell("LocPinX", 0), Cell("LocPinY", 0), Cell("Angle", 0)),
            new XElement(Legacy + "Line", Cell("LineWeight", .05), new XElement(Legacy + "LineColor", "#0000FF")),
            new XElement(Legacy + "Geom", new XAttribute("IX", "0"), Cell("NoFill", 1), Cell("NoLine", 0), Cell("NoShow", 0),
                new XElement(Legacy + "MoveTo", new XAttribute("IX", "5"), Cell("X", startX), Cell("Y", startY)), arc));
        VisioDocument document = LoadXml(Encoding.UTF8.GetBytes(new XElement(Legacy + "VisioDocument", new XElement(Legacy + "Masters",
            new XElement(Legacy + "Master", new XAttribute("ID", "1"), new XAttribute("NameU", "Arc"), new XElement(Legacy + "Shapes", shape)))).ToString()));
        if (!relative) return document;
        // Construct a native Open XML fixture, rather than inventing relative
        // elements in the older schema or adding a product hook for tests.
        document.AddPage("Seed").AddShape("seed", "Arc", 2, 6, 4, 4);
        return RewriteMaster(document, native => {
            foreach (XElement row in native.Descendants(Modern + "Row")) {
                row.SetAttributeValue("T", "Rel" + (string)row.Attribute("T")!);
                foreach (XElement cell in row.Elements(Modern + "Cell").Where(c => (string?)c.Attribute("N") is "X" or "Y" or "A" or "B"))
                    cell.SetAttributeValue("V", (double.Parse((string)cell.Attribute("V")!, CultureInfo.InvariantCulture) / 4).ToString("R", CultureInfo.InvariantCulture));
            }
        });
    }

    private static VisioDocument RewriteMaster(VisioDocument document, Action<XDocument> edit) {
        using var bytes = new MemoryStream();
        byte[] package = document.ToBytes();
        bytes.Write(package, 0, package.Length);
        using (var zip = new ZipArchive(bytes, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry part = zip.GetEntry("visio/masters/master1.xml")!;
            XDocument native;
            using (var stream = part.Open()) native = XDocument.Load(stream);
            edit(native);
            part.Delete();
            using var replacement = zip.CreateEntry("visio/masters/master1.xml").Open();
            native.Save(replacement);
        }
        bytes.Position = 0;
        return VisioDocument.Load(bytes);
    }

    private static XElement Cell(string name, double value) => new(Legacy + name, value.ToString("R", CultureInfo.InvariantCulture));
}
