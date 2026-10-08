using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioNativeReflectionRenderingTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";

    // The asymmetric outline coordinates are independently read by libvisio 0.1.11.
    // Parent and leaf reflections must compose about LocPin before each rotation.
    [Theory]
    [InlineData(false, false, 0, false, 0, "122.4 374.4 266.4 374.4 151.2 302.4")]
    [InlineData(true, false, 0, false, 0, "165.6 374.4 21.6 374.4 136.8 302.4")]
    [InlineData(false, true, 0, false, 0, "122.4 345.6 266.4 345.6 151.2 417.6")]
    [InlineData(true, false, 30, false, 0, "169.9061 361.6708 45.1985 433.6708 108.9646 313.7169")]
    [InlineData(false, false, 0, true, 0, "165.6 158.4 21.6 158.4 136.8 86.4")]
    [InlineData(false, true, 30, true, 45, "50.5786 289.5885 -88.5148 326.8584 41.3949 366.5891")]
    public void ReflectionsUseLocalPinsAndContainingGroupTransformsInNativeAndDrawingOutlines(
        bool flipX, bool flipY, double angle, bool parentX, double parentAngle, string coordinates) {
        VisioDocument document = Load(Create(flipX, flipY, angle, parentX, parentAngle));
        AssertOutlines(document, coordinates);
        AssertOutlines(VisioDocument.Load(new MemoryStream(document.ToBytes())), coordinates);
        AssertOutlines(VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value, coordinates);
    }

    [Fact]
    public void InheritedMasterReflectionAndExplicitZeroOverridesSurviveReopening() {
        XDocument xml = Create(true, false);
        XElement shape = xml.Descendants(Legacy + "Shape").Single();
        xml.Root!.AddFirst(new XElement(Legacy + "Masters", new XElement(Legacy + "Master", new XAttribute("ID", "0"),
            new XAttribute("NameU", "Reflected"), new XElement(Legacy + "Shapes", new XElement(shape)))));
        shape.SetAttributeValue("Master", "0");
        XElement reflection = shape.Element(Legacy + "XForm")!.Element(Legacy + "FlipX")!;
        reflection.SetAttributeValue("F", "Inh"); reflection.Value = "0";
        VisioDocument inherited = Load(xml);
        AssertOutlines(inherited, "165.6 374.4 21.6 374.4 136.8 302.4");
        AssertOutlines(VisioDocument.Load(new MemoryStream(inherited.ToBytes())), "165.6 374.4 21.6 374.4 136.8 302.4");
        reflection.Attribute("F")!.Remove();
        VisioDocument overridden = Load(xml);
        AssertOutlines(overridden, "122.4 374.4 266.4 374.4 151.2 302.4");
    }

    [Fact]
    public void CachedFormulaLossAndUnusableReflectionsAreReportedAcrossNativeAndDrawingRoutes() {
        XDocument xml = Create(true, false);
        XElement flip = xml.Descendants(Legacy + "FlipX").Single();
        flip.SetAttributeValue("F", "Sheet.99!FlipX");
        VisioDocument document = Load(xml);
        AssertOutlines(document, "165.6 374.4 21.6 374.4 136.8 302.4");
        var scene = document.ToDrawings();
        Assert.Contains(scene.Report.FidelityDiagnostics, item => item.Code == "VISIO_SHAPE_CACHED_REFLECTION" && item.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Throws<OfficeConversionException>(() => scene.Report.RequireNoLoss());
        foreach (OfficeImageExportFormat format in new[] { OfficeImageExportFormat.Svg, OfficeImageExportFormat.Png }) {
            var result = document.ExportImage(format, new VisioImageExportOptions { RenderText = false, RenderStencilArtwork = false });
            Assert.Contains(result.Diagnostics, item => item.Code == "VISIO_SHAPE_CACHED_REFLECTION");
        }
        flip.Value = "NaN"; document = Load(xml);
        Assert.Contains(document.ToDrawings().Report.FidelityDiagnostics, item => item.Code == "VISIO_SHAPE_TRANSFORM_INVALID" && item.LossKind == OfficeConversionLossKind.Omission);
        foreach (OfficeImageExportFormat format in new[] { OfficeImageExportFormat.Svg, OfficeImageExportFormat.Png }) {
            var result = document.ExportImage(format, new VisioImageExportOptions { RenderText = false, RenderStencilArtwork = false });
            Assert.Contains(result.Diagnostics, item => item.Code == "VISIO_SHAPE_TRANSFORM_INVALID" && item.LossKind == OfficeConversionLossKind.Omission);
        }
    }

    [Fact]
    public void ReflectedNativeAndDrawingRasterPaintTheSameVisibleAsymmetricTriangle() {
        VisioDocument document = Load(Create(true, false)); VisioPage page = document.Pages[0];
        var native = VisualBaselineTestSupport.DecodePng(page.ToPng(new VisioPngSaveOptions {
            PixelsPerInch = 72, Supersampling = 1, RenderText = false, RenderStencilArtwork = false
        }), "Native reflected PNG");
        var drawing = VisualBaselineTestSupport.DecodePng(OfficeDrawingRasterRenderer.ToPng(page.ToDrawing().Value,
            background: OfficeColor.White), "Drawing reflected PNG");
        foreach (OfficeRasterImage image in new[] { native, drawing }) {
            Assert.Equal(OfficeColor.Red, image.GetPixel(140, 345));
            Assert.Equal(OfficeColor.White, image.GetPixel(210, 345));
        }
    }

    private static void AssertOutlines(VisioDocument document, string coordinates) {
        byte[] before = document.ToLegacyXmlResult().Value;
        double[] expected = Numbers(coordinates);
        string[] outputs = {
            document.Pages[0].ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 72, RenderText = false, RenderStencilArtwork = false }),
            OfficeDrawingSvgExporter.ToSvg(document.Pages[0].ToDrawing().Value, 1, OfficeSvgSizeUnit.Point)
        };
        foreach (string output in outputs) {
            XElement path = XDocument.Parse(output).Descendants(Svg + "path").First(item =>
                string.Equals((string?)item.Attribute("fill"), "#FF0000", StringComparison.OrdinalIgnoreCase));
            double[] actual = Numbers(path.Attribute("d")!.Value);
            // Off-page scenes place local path coordinates in translated effect groups.
            // Compare the painted page coordinates rather than incidental local bounds.
            OfficeTransform placement = OfficeTransform.Identity;
            foreach (XElement ancestor in path.Ancestors().Where(item => item.Attribute("transform") != null)) {
                string transform = ancestor.Attribute("transform")!.Value;
                Assert.StartsWith("matrix(", transform);
                double[] matrix = Numbers(transform); Assert.Equal(6, matrix.Length);
                placement = placement.Then(new OfficeTransform(matrix[0], matrix[1], matrix[2], matrix[3], matrix[4], matrix[5]));
            }
            for (int index = 0; index + 1 < actual.Length; index += 2) {
                OfficePoint point = placement.TransformPoint(new OfficePoint(actual[index], actual[index + 1]));
                actual[index] = point.X; actual[index + 1] = point.Y;
            }
            Assert.True(actual.Length >= expected.Length);
            for (int index = 0; index < expected.Length; index++) Assert.InRange(actual[index], expected[index] - 0.002, expected[index] + 0.002);
        }
        Assert.Equal(before, document.ToLegacyXmlResult().Value);
    }

    private static double[] Numbers(string value) => Regex.Matches(value, @"-?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?")
        .Cast<Match>().Select(match => double.Parse(match.Value, CultureInfo.InvariantCulture)).ToArray();

    private static VisioDocument Load(XDocument xml) => VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml.ToString()))).Value;

    private static XDocument Create(bool flipX, bool flipY, double angle = 0, bool parentX = false, double parentAngle = 0) {
        string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);
        string shape = $"""
            <Shape ID="1" NameU="Asymmetric" Type="Shape"><XForm><PinX>2</PinX><PinY>3</PinY><Width>2</Width><Height>1</Height>
            <LocPinX>0.3</LocPinX><LocPinY>0.2</LocPinY><Angle>{Number(angle * Math.PI / 180)}</Angle><FlipX>{(flipX ? 1 : 0)}</FlipX><FlipY>{(flipY ? 1 : 0)}</FlipY></XForm>
            <Line><LinePattern>0</LinePattern></Line><Fill><FillForegnd>#FF0000</FillForegnd><FillPattern>1</FillPattern></Fill>
            <Geom IX="0"><MoveTo IX="0"><X>0</X><Y>0</Y></MoveTo><LineTo IX="1"><X>2</X><Y>0</Y></LineTo>
            <LineTo IX="2"><X>0.4</X><Y>1</Y></LineTo><LineTo IX="3"><X>0</X><Y>0</Y></LineTo></Geom></Shape>
            """;
        if (parentX) shape = $"<Shape ID='2' Type='Group'><XForm><PinX>4</PinX><PinY>3</PinY><Width>4</Width><Height>4</Height><LocPinX>0</LocPinX><LocPinY>0</LocPinY><FlipX>1</FlipX><Angle>{Number(parentAngle * Math.PI / 180)}</Angle></XForm><Fill><FillPattern>0</FillPattern></Fill><Line><LinePattern>0</LinePattern></Line><Shapes>{shape}</Shapes></Shape>";
        return XDocument.Parse($"<VisioDocument xmlns='{Legacy}'><Pages><Page ID='0' Name='Reflected'><PageSheet><PageProps><PageWidth>8</PageWidth><PageHeight>8</PageHeight></PageProps></PageSheet><Shapes>{shape}</Shapes></Page></Pages></VisioDocument>");
    }
}
