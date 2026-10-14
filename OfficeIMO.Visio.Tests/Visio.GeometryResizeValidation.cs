using System.IO.Compression;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioGeometryResizeValidationTests {
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";

    // Cacheless cross-sheet formulas are preserved on load, but cannot be
    // recalculated by the supported resize profile. Check distinct row roles,
    // including curve controls and profile cells that uniform scaling retains.
    [Theory]
    [InlineData("MoveTo", "Y")]
    [InlineData("LineTo", "X")]
    [InlineData("RelLineTo", "X")]
    [InlineData("ArcTo", "A")]
    [InlineData("EllipticalArcTo", "C")]
    [InlineData("RelEllipticalArcTo", "D")]
    [InlineData("Ellipse", "C")]
    [InlineData("InfiniteLine", "B")]
    [InlineData("CubBezTo", "D")]
    [InlineData("RelCubBezTo", "D")]
    [InlineData("QuadBezTo", "A")]
    [InlineData("RelQuadBezTo", "B")]
    [InlineData("PolylineTo", "X")]
    [InlineData("NURBSTo", "D")]
    [InlineData("SplineStart", "A")]
    [InlineData("SplineKnot", "A")]
    public void CachelessCoordinatesRejectExistingResizeAndMasterCreationBeforeMutation(string type, string cellName) {
        foreach (bool creation in new[] { false, true }) {
            var document = Load(type, cellName);
            var page = document.Pages[0]; var shape = page.Shapes[0];
            VisioMaster? master = creation ? document.RegisterMaster("Unreadable", shape, "9") : null;
            if (creation) page = document.AddPage("Destination");
            string before = NativeParts(document);
            Assert.Throws<NotSupportedException>(() => {
                if (creation) page.AddShape("instance", master!, 3, 3, 2, 2);
                else page.ResizeShape(shape, 2, 2);
            });
            Assert.Equal(before, NativeParts(document));
            Assert.Equal(1, shape.Width); Assert.Equal(1, shape.Height);
            if (creation) Assert.Empty(page.Shapes);
        }
    }

    [Theory]
    [InlineData("LineTo", "X")]
    [InlineData("RelCubBezTo", "D")]
    [InlineData("NURBSTo", "E")]
    [InlineData("PolylineTo", "A")]
    public void IncompleteNativeCoordinateRowsRejectBeforeMutation(string type, string cellName) {
        var document = Load(type, cellName, missing: true); var page = document.Pages[0];
        string before = NativeParts(document);
        Assert.Throws<NotSupportedException>(() => page.ResizeShape(page.Shapes[0], 2, 3));
        Assert.Equal(before, NativeParts(document));
    }

    [Fact]
    public void FiniteCacheAllowsUnsupportedFormulaToBeMaterializedAndProducerMetadataToSurvive() {
        var document = Load("LineTo", "X", cache: "0.5"); var page = document.Pages[0];
        page.ResizeShape(page.Shapes[0], 2, 3);
        var reopened = VisioDocument.Load(new MemoryStream(document.ToBytes()));
        XElement row = PageXml(reopened).Descendants(Modern + "Row").Single(r => (string?)r.Attribute("IX") == "2");
        XElement coordinate = row.Elements(Modern + "Cell").Single(c => (string?)c.Attribute("N") == "X");
        Assert.Equal("1", (string?)coordinate.Attribute("V")); Assert.Null(coordinate.Attribute("F"));
        XElement metadata = row.Elements(Modern + "Cell").Single(c => (string?)c.Attribute("N") == "ProducerTag");
        Assert.Equal("kept", (string?)metadata.Attribute("V")); Assert.Equal("Sheet.2!User.Tag", (string?)metadata.Attribute("F"));
    }

    [Theory]
    [InlineData(0)] // Last knot.
    [InlineData(1)] // Degree.
    [InlineData(6)] // Control knot.
    [InlineData(7)] // Control weight.
    public void CachelessNurbsScalarArgumentsRejectResizeAndCreationBeforeMutation(int argument) {
        var values = new[] { "2", "2", "1", "1", "0.5", "0.8", "1", "1" };
        values[argument] = "Sheet.2!Width";
        foreach (bool creation in new[] { false, true }) {
            var document = Load("NURBSTo", "X", cache: "1", controlFormula: "NURBS(" + string.Join(",", values) + ")");
            var page = document.Pages[0]; var shape = page.Shapes[0];
            VisioMaster? master = creation ? document.RegisterMaster("Unreadable", shape, "9") : null;
            if (creation) page = document.AddPage("Destination");
            string before = NativeParts(document);
            Assert.Throws<NotSupportedException>(() => {
                if (creation) page.AddShape("instance", master!, 3, 3, 2, 2);
                else page.ResizeShape(shape, 2, 2);
            });
            Assert.Equal(before, NativeParts(document));
            if (creation) Assert.Empty(page.Shapes);
        }
    }

    [Fact]
    public void EvaluableNurbsScalarArgumentsKeepTheOriginalRationalCurveDuringResize() {
        var document = Load("NURBSTo", "X", cache: "1", controlFormula: "NURBS(Width*2,Width+1,1,1,0.5,0.8,Width,Width)");
        var page = document.Pages[0]; double[] before = PaintedCoordinates(page);
        double centerX = 2 * 40, centerY = (page.Height - 2) * 40;
        page.ResizeShape(page.Shapes[0], 2, 2);
        foreach (var candidate in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())),
                     VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value }) {
            double[] after = PaintedCoordinates(candidate.Pages[0]); Assert.Equal(before.Length, after.Length);
            for (int i = 0; i < before.Length; i++) {
                double center = i % 2 == 0 ? centerX : centerY;
                Assert.InRange(after[i], center + 2 * (before[i] - center) - .002, center + 2 * (before[i] - center) + .002);
            }
            XElement formula = PageXml(candidate).Descendants(Modern + "Row").Single(r => (string?)r.Attribute("T") == "NURBSTo")
                .Elements(Modern + "Cell").Single(c => (string?)c.Attribute("N") == "E");
            Assert.Equal("NURBS(2,2,1,1,1,1.6,1,1)", (string?)formula.Attribute("F"));
        }
    }

    private static VisioDocument Load(string type, string cellName, bool missing = false, string? cache = null, string? controlFormula = null) {
        var document = VisioDocument.Create(); var page = document.AddPage("Source"); page.AddRectangle(2, 2, 1, 1);
        using var bytes = new MemoryStream();
        byte[] packageBytes = document.ToBytes();
        bytes.Write(packageBytes, 0, packageBytes.Length);
        using (var zip = new ZipArchive(bytes, ZipArchiveMode.Update, true)) {
            var entry = zip.GetEntry("visio/pages/page1.xml")!; XDocument xml;
            using (var stream = entry.Open()) xml = XDocument.Load(stream);
            XElement section = xml.Descendants(Modern + "Section").Single(s => (string?)s.Attribute("N") == "Geometry");
            section.Elements(Modern + "Row").Remove();
            section.Add(new XElement(Modern + "Row", new XAttribute("T", "MoveTo"), new XAttribute("IX", "1"), Cell("X", "0"), Cell("Y", "0")));
            XElement row = new(Modern + "Row", new XAttribute("T", type), new XAttribute("IX", "2"),
                Cell("X", "1"), Cell("Y", "1"), Cell("ProducerTag", "kept", "Sheet.2!User.Tag"));
            foreach (char name in "ABCD") row.Add(Cell(name.ToString(), "0.5"));
            if (type == "PolylineTo") row.Elements(Modern + "Cell").Single(c => (string?)c.Attribute("N") == "A").SetAttributeValue("F", "POLYLINE(1,1,0,0,1,1)");
            if (type == "NURBSTo") {
                foreach (var value in new[] { (Name: "A", Value: "1"), (Name: "B", Value: "1"), (Name: "C", Value: "0"), (Name: "D", Value: "1") })
                    row.Elements(Modern + "Cell").Single(c => (string?)c.Attribute("N") == value.Name).SetAttributeValue("V", value.Value);
                row.Add(Cell("E", controlFormula ?? "NURBS(1,1,1,1,0,0,0,1,1,1,1,1)"));
            }
            XElement target = row.Elements(Modern + "Cell").Single(c => (string?)c.Attribute("N") == cellName);
            if (missing) target.Remove();
            else { target.SetAttributeValue("V", cache); target.SetAttributeValue("F", "Sheet.2!Width"); }
            section.Add(row);
            entry.Delete(); using var output = zip.CreateEntry("visio/pages/page1.xml").Open(); xml.Save(output);
        }
        bytes.Position = 0; return VisioDocument.Load(bytes);
    }

    private static XElement Cell(string name, string value, string? formula = null) =>
        new(Modern + "Cell", new XAttribute("N", name), new XAttribute("V", value), formula == null ? null : new XAttribute("F", formula));
    private static double[] PaintedCoordinates(VisioPage page) {
        string data = XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 40, RenderText = false, RenderStencilArtwork = false }))
            .Descendants().Single(e => e.Name.LocalName == "path" && (string?)e.Attribute("data-officeimo-preserved-geometry") == "true").Attribute("d")!.Value;
        return Regex.Matches(data, @"[-+]?(?:\d*\.)?\d+(?:[Ee][-+]?\d+)?").Cast<Match>()
            .Select(m => double.Parse(m.Value, CultureInfo.InvariantCulture)).ToArray();
    }
    private static XElement PageXml(VisioDocument document) {
        using var zip = new ZipArchive(new MemoryStream(document.ToBytes())); using var stream = zip.GetEntry("visio/pages/page1.xml")!.Open();
        return XDocument.Load(stream).Root!;
    }
    private static string NativeParts(VisioDocument document) {
        using var zip = new ZipArchive(new MemoryStream(document.ToBytes()));
        return string.Join("\n", zip.Entries.OrderBy(e => e.FullName).Select(e => {
            using var stream = e.Open(); using var copy = new MemoryStream(); stream.CopyTo(copy);
            return e.FullName + ":" + Convert.ToBase64String(copy.ToArray());
        }));
    }
}
