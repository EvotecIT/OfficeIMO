using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Reader.Visio;
using OfficeIMO.Visio.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioLegacyBinaryTests {
    private static string Fixture(string extension) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyBinary", "pronom-visio2003." + extension);
    private static IEnumerable<VisioShape> Flatten(IEnumerable<VisioShape> shapes) => shapes.SelectMany(shape => new[] { shape }.Concat(Flatten(shape.Children)));

    [Theory]
    [InlineData("vsd", "vdx", VisioPackageType.Drawing)]
    [InlineData("vst", "vtx", VisioPackageType.Template)]
    [InlineData("vss", "vsx", VisioPackageType.Stencil)]
    public void ReadsIndependentNativeFamiliesThroughTheExistingModel(string binary, string xml, VisioPackageType family) {
        var imported = VisioDocument.LoadLegacyBinary(Fixture(binary));
        VisioDocument expected = VisioDocument.LoadLegacyXml(Fixture(xml)).Value;
        Assert.Equal(family, imported.Value.PackageType);
        Assert.Equal(expected.Pages.Count, imported.Value.Pages.Count);
        Assert.Equal(3, imported.Value.Masters.Count);
        Assert.Equal(expected.Masters.Select(master => master.Id), imported.Value.Masters.Select(master => master.Id));
        Assert.All(imported.Value.Masters, master => Assert.Equal("Master-" + master.Id, master.NameU));
        Assert.Null(imported.Value.FilePath);
        Assert.Equal(OfficeLegacyImportQuality.Structured, imported.Report.Quality);
        Assert.True(imported.Report.HasLoss);
        Assert.Throws<InvalidOperationException>(() => imported.Report.RequireNoLoss());
        for (int index = 0; index < expected.Pages.Count; index++) {
            VisioShape[] shapes = Flatten(imported.Value.Pages[index].Shapes).ToArray();
            VisioShape[] originals = Flatten(expected.Pages[index].Shapes).ToArray();
            Assert.Equal(33, shapes.Length);
            Assert.Equal(originals.Select(shape => shape.Id), shapes.Select(shape => shape.Id));
            for (int shapeIndex = 0; shapeIndex < shapes.Length; shapeIndex++) {
                Assert.Equal(originals[shapeIndex].PinX, shapes[shapeIndex].PinX, 8);
                Assert.Equal(originals[shapeIndex].PinY, shapes[shapeIndex].PinY, 8);
                Assert.Equal(originals[shapeIndex].Width, shapes[shapeIndex].Width, 8);
                Assert.Equal(originals[shapeIndex].Height, shapes[shapeIndex].Height, 8);
            }
            Assert.Equal(expected.Pages[index].Width, imported.Value.Pages[index].Width, 8);
            Assert.Equal(expected.Pages[index].Height, imported.Value.Pages[index].Height, 8);
            Assert.Contains(shapes, shape => (shape.Text ?? "").Contains("Office"));
            Assert.Contains(shapes, shape => (shape.Text ?? "").Contains("80 sq. ft."));
        }
        using var package = new MemoryStream();
        imported.Value.Save(package);
        VisioDocument reopened = VisioDocument.Load(package);
        Assert.Equal(family, reopened.PackageType);
        Assert.Equal(imported.Value.Pages.Count, reopened.Pages.Count);
        Assert.Equal(3, reopened.Masters.Count);
    }

    [Fact]
    public void NativeDrawingUsesTheExistingSvgAndPdfConversionRoutes() {
        var imported = VisioDocument.LoadLegacyBinary(Fixture("vsd"));
        string svg = Assert.Single(imported.Value.Pages).ToSvg();
        XDocument artifact = XDocument.Parse(svg);
        Assert.Equal("svg", artifact.Root!.Name.LocalName);
        Assert.Contains("Office", svg);
        XNamespace ns = artifact.Root.Name.Namespace;
        foreach (var pair in new[] { ("30", "31", 5), ("32", "33", 7) }) {
            XElement outline = Assert.Single(artifact.Descendants(ns + "g"), element => (string?)element.Attribute("data-visio-shape-id") == pair.Item1);
            Assert.Contains(outline.Elements(ns + "path"), path => ((string?)path.Attribute("d") ?? "").Count(character => character == 'L') >= pair.Item3);
            XElement text = Assert.Single(artifact.Descendants(ns + "g"), element => (string?)element.Attribute("data-visio-shape-id") == pair.Item2);
            Assert.Empty(text.Elements(ns + "path"));
            VisioShape shape = Flatten(imported.Value.Pages[0].Shapes).Single(shape => shape.Id == pair.Item1);
            Assert.Equal(2, shape.FillPattern);
            Assert.Equal(64, shape.FillColor.A);
        }
        var drawing = imported.Value.Pages[0].ToDrawing();
        Assert.DoesNotContain(drawing.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "VISIO_DRAWING_GEOMETRY_FALLBACK");
        Assert.Contains(drawing.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "VISIO_DRAWING_FILL_PATTERN");
        var conversion = imported.Value.ToPdfDocumentResult(new VisioToPdfOptions { Mode = VisioPdfProjectionMode.DiagramPages });
        byte[] pdf = conversion.ToBytes();
        Assert.True(pdf.Length > 100);
        Assert.Equal("%PDF-", Encoding.ASCII.GetString(pdf, 0, 5));
    }

    [Fact]
    public void PairedNativeXmlKeepsTextOnlyMasterChildrenAndCachedPaint() {
        var document = VisioDocument.LoadLegacyXml(Fixture("vdx")).Value;
        XDocument svg = XDocument.Parse(document.Pages[0].ToSvg());
        XNamespace ns = svg.Root!.Name.Namespace;
        XElement text = Assert.Single(svg.Descendants(ns + "g"), element => (string?)element.Attribute("data-visio-shape-id") == "31");
        Assert.Empty(text.Elements(ns + "path"));
        VisioShape outline = Flatten(document.Pages[0].Shapes).Single(shape => shape.Id == "30");
        Assert.Equal(2, outline.FillPattern);
        Assert.Equal(64, outline.FillColor.A);
    }

    [Fact]
    public void LocalTransparencyOverridesAnInheritedMasterColor() {
        const string xml = """
            <VisioDocument xmlns="http://schemas.microsoft.com/visio/2003/core">
              <Masters><Master ID="1" NameU="Test"><Shapes><Shape ID="1">
                <XForm><Width>2</Width><Height>2</Height><LocPinX>1</LocPinX><LocPinY>1</LocPinY></XForm>
                <Fill><FillForegnd>#008000</FillForegnd><FillForegndTrans>0.75</FillForegndTrans></Fill>
                <Line><LineColor>#0000FF</LineColor><LineColorTrans>0.75</LineColorTrans></Line>
              </Shape></Shapes></Master></Masters>
              <Pages><Page ID="0" NameU="Page"><PageSheet><PageProps><PageWidth>10</PageWidth><PageHeight>10</PageHeight></PageProps></PageSheet>
                <Shapes><Shape ID="1" Master="1"><XForm><PinX>3</PinX><PinY>3</PinY></XForm>
                  <Fill><FillForegndTrans>0.5</FillForegndTrans></Fill>
                  <Line><LineColorTrans>0.5</LineColorTrans></Line>
                </Shape></Shapes>
              </Page></Pages>
            </VisioDocument>
            """;
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        var document = VisioDocument.LoadLegacyXml(input).Value;
        var shape = Assert.Single(Assert.Single(document.Pages).Shapes);
        Assert.Equal(0, shape.FillColor.R);
        Assert.Equal(128, shape.FillColor.G);
        Assert.Equal(128, shape.FillColor.A);
        Assert.Equal(255, shape.LineColor.B); Assert.Equal(128, shape.LineColor.A);
    }

    [Fact]
    public void LocalRgbOverridesKeepIndependentMasterTransparency() {
        const string xml = """
            <VisioDocument xmlns="http://schemas.microsoft.com/visio/2003/core">
              <Masters><Master ID="1" NameU="Test"><Shapes><Shape ID="1">
                <XForm><Width>2</Width><Height>2</Height></XForm>
                <Fill><FillForegnd>#008000</FillForegnd><FillForegndTrans>0.75</FillForegndTrans></Fill>
                <Line><LineColor>#0000FF</LineColor><LineColorTrans>0.5</LineColorTrans></Line>
              </Shape></Shapes></Master></Masters>
              <Pages><Page ID="0"><PageSheet><PageProps><PageWidth>10</PageWidth><PageHeight>10</PageHeight></PageProps></PageSheet>
                <Shapes><Shape ID="1" Master="1"><XForm><PinX>3</PinX><PinY>3</PinY></XForm>
                  <Fill><FillForegnd>#FF0000</FillForegnd></Fill><Line><LineColor>#00FFFF</LineColor></Line>
                </Shape></Shapes>
              </Page></Pages>
            </VisioDocument>
            """;
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        var document = VisioDocument.LoadLegacyXml(input).Value;
        var shape = Assert.Single(Assert.Single(document.Pages).Shapes);
        Assert.Equal(255, shape.FillColor.R); Assert.Equal(0, shape.FillColor.G); Assert.Equal(64, shape.FillColor.A);
        Assert.Equal(0, shape.LineColor.R); Assert.Equal(255, shape.LineColor.G); Assert.Equal(128, shape.LineColor.A);
        using var package = new MemoryStream(); document.Save(package);
        var reopened = Assert.Single(Assert.Single(VisioDocument.Load(package).Pages).Shapes);
        Assert.Equal(shape.FillColor, reopened.FillColor); Assert.Equal(shape.LineColor, reopened.LineColor);
    }

    [Fact]
    public void TopLevelConnectorsResolveCachedStyleAndLocalPaint() {
        const string xml = """
            <VisioDocument xmlns="http://schemas.microsoft.com/visio/2003/core">
              <StyleSheets><StyleSheet ID="2"><Line><LineColor>#0000FF</LineColor><LineColorTrans>0.5</LineColorTrans>
                <LineWeight>0.1</LineWeight><LinePattern>2</LinePattern></Line></StyleSheet></StyleSheets>
              <Pages><Page ID="0"><PageSheet><PageProps><PageWidth>10</PageWidth><PageHeight>10</PageHeight></PageProps></PageSheet>
                <Shapes><Shape ID="1" LineStyle="2"><XForm1D><BeginX>1</BeginX><BeginY>1</BeginY><EndX>4</EndX><EndY>4</EndY></XForm1D></Shape>
                  <Shape ID="2" LineStyle="2"><XForm1D><BeginX>1</BeginX><BeginY>5</BeginY><EndX>4</EndX><EndY>8</EndY></XForm1D>
                    <Line><LineColor>#FF0000</LineColor><LineColorTrans>0.75</LineColorTrans></Line>
                  </Shape></Shapes>
              </Page></Pages>
            </VisioDocument>
            """;
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        var document = VisioDocument.LoadLegacyXml(input).Value;
        var page = Assert.Single(document.Pages);
        Assert.Equal(2, page.Connectors.Count);
        Assert.Equal(255, page.Connectors[0].LineColor.B); Assert.Equal(128, page.Connectors[0].LineColor.A);
        Assert.Equal(0.1, page.Connectors[0].LineWeight); Assert.Equal(2, page.Connectors[0].LinePattern);
        Assert.Equal(255, page.Connectors[1].LineColor.R); Assert.Equal(64, page.Connectors[1].LineColor.A);
        Assert.Contains("stroke-opacity", page.ToSvg());
        using var package = new MemoryStream(); document.Save(package);
        var reopened = Assert.Single(VisioDocument.Load(package).Pages);
        Assert.Equal(page.Connectors.Select(connector => connector.LineColor), reopened.Connectors.Select(connector => connector.LineColor));
        page.Connectors[1].LineColor = OfficeIMO.Drawing.OfficeColor.FromRgba(255, 0, 0, 255);
        using var edited = new MemoryStream(); document.Save(edited);
        Assert.Equal(255, Assert.Single(VisioDocument.Load(edited).Pages).Connectors[1].LineColor.A);
        edited.Position = 0;
        using var archive = new System.IO.Compression.ZipArchive(edited, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: true);
        using var pageXml = archive.GetEntry("visio/pages/page1.xml")!.Open();
        var saved = XDocument.Load(pageXml);
        XNamespace ns = saved.Root!.Name.Namespace;
        Assert.All(saved.Descendants(ns + "Shape"), shape => Assert.Single(shape.Elements(ns + "Cell"),
            cell => (string?)cell.Attribute("N") == "LineColorTrans"));
    }

    [Fact]
    public void MasterBackedConnectorsResolveEachPaintPropertyBeforeOverrides() {
        const string xml = """
            <VisioDocument xmlns="http://schemas.microsoft.com/visio/2003/core">
              <StyleSheets><StyleSheet ID="2"><Line><LineColor>#008000</LineColor><LineColorTrans>0.25</LineColorTrans>
                <LineWeight>0.2</LineWeight><LinePattern>1</LinePattern></Line></StyleSheet></StyleSheets>
              <Masters><Master ID="1" NameU="Dynamic connector"><Shapes><Shape ID="10">
                <XForm1D><BeginX>0</BeginX><BeginY>0</BeginY><EndX>1</EndX><EndY>1</EndY></XForm1D>
                <Line><LineColor>#0000FF</LineColor><LineColorTrans>0.5</LineColorTrans><LineWeight>0.1</LineWeight><LinePattern>2</LinePattern></Line>
                <Shapes><Shape ID="11"><XForm1D><BeginX>0</BeginX><BeginY>0</BeginY><EndX>1</EndX><EndY>1</EndY></XForm1D>
                  <Line><LineColor>#FF0000</LineColor><LineColorTrans>0.25</LineColorTrans><LineWeight>0.3</LineWeight><LinePattern>3</LinePattern></Line>
                </Shape></Shapes>
              </Shape></Shapes></Master></Masters>
              <Pages><Page ID="0"><PageSheet><PageProps><PageWidth>10</PageWidth><PageHeight>10</PageHeight></PageProps></PageSheet>
                <Shapes><Shape ID="1" Master="1"><XForm1D><BeginX>1</BeginX><BeginY>1</BeginY><EndX>4</EndX><EndY>1</EndY></XForm1D></Shape>
                  <Shape ID="2" Master="1"><XForm1D><BeginX>1</BeginX><BeginY>2</BeginY><EndX>4</EndX><EndY>2</EndY></XForm1D><Line><LineColor>#FF0000</LineColor></Line></Shape>
                  <Shape ID="3" Master="1"><XForm1D><BeginX>1</BeginX><BeginY>3</BeginY><EndX>4</EndX><EndY>3</EndY></XForm1D><Line><LineColorTrans>0.75</LineColorTrans></Line></Shape>
                  <Shape ID="4" Master="1" LineStyle="2"><XForm1D><BeginX>1</BeginX><BeginY>4</BeginY><EndX>4</EndX><EndY>4</EndY></XForm1D></Shape>
                  <Shape ID="5" Master="1" MasterShape="11"><XForm1D><BeginX>1</BeginX><BeginY>5</BeginY><EndX>4</EndX><EndY>5</EndY></XForm1D></Shape>
                </Shapes>
              </Page></Pages>
            </VisioDocument>
            """;
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        var document = VisioDocument.LoadLegacyXml(input).Value;
        var page = Assert.Single(document.Pages);
        Assert.Equal(5, page.Connectors.Count);
        Assert.Equal(255, page.Connectors[0].LineColor.B); Assert.Equal(128, page.Connectors[0].LineColor.A);
        Assert.Equal(255, page.Connectors[1].LineColor.R); Assert.Equal(128, page.Connectors[1].LineColor.A);
        Assert.Equal(255, page.Connectors[2].LineColor.B); Assert.Equal(64, page.Connectors[2].LineColor.A);
        Assert.Equal(128, page.Connectors[3].LineColor.G); Assert.Equal(191, page.Connectors[3].LineColor.A);
        Assert.All(page.Connectors.Take(3), connector => { Assert.Equal(0.1, connector.LineWeight); Assert.Equal(2, connector.LinePattern); });
        Assert.Equal(0.2, page.Connectors[3].LineWeight); Assert.Equal(1, page.Connectors[3].LinePattern);
        Assert.Equal(255, page.Connectors[4].LineColor.R); Assert.Equal(191, page.Connectors[4].LineColor.A);
        Assert.Equal(0.3, page.Connectors[4].LineWeight); Assert.Equal(3, page.Connectors[4].LinePattern);
        using var package = new MemoryStream(); document.Save(package);
        var reopened = Assert.Single(VisioDocument.Load(package).Pages);
        Assert.Equal(page.Connectors.Select(connector => connector.LineColor), reopened.Connectors.Select(connector => connector.LineColor));
    }

    [Fact]
    public void CallerStreamIsRestoredAndLimitsAndCancellationAreEnforced() {
        using var source = File.OpenRead(Fixture("vsd")); source.Position = 9;
        VisioDocument.LoadLegacyBinary(source);
        Assert.True(source.CanRead); Assert.Equal(9, source.Position);
        Assert.Throws<InvalidDataException>(() => VisioDocument.LoadLegacyBinary(source, options: new() { Limits = new() { MaxInputBytes = 128 } }));
        Assert.Throws<InvalidDataException>(() => VisioDocument.LoadLegacyBinary(source, options: new() { MaxDecompressedBytes = 128 }));
        Assert.Throws<InvalidDataException>(() => VisioDocument.LoadLegacyBinary(source, options: new() { Limits = new() { MaxItems = 1 } }));
        Assert.Throws<InvalidDataException>(() => VisioDocument.LoadLegacyBinary(source, options: new() { Limits = new() { MaxTextCharacters = 1 } }));
        Assert.Throws<InvalidDataException>(() => VisioDocument.LoadLegacyBinary(source, options: new() { Limits = new() { MaxRecords = 1 } }));
        Assert.Throws<InvalidDataException>(() => VisioDocument.LoadLegacyBinary(source, options: new() { MaxDepth = 1 }));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => VisioDocument.LoadLegacyBinary(source, cancellationToken: cancelled.Token));
        Assert.Equal(9, source.Position); Assert.True(source.CanRead);
    }

    [Fact]
    public void UnsupportedGenerationsAndInvalidContainersFailExplicitly() {
        string old = Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyBinary", "pronom-visio-v5.vsd");
        Assert.Throws<NotSupportedException>(() => VisioDocument.LoadLegacyBinary(old));
        using var source = new MemoryStream(new byte[512]);
        Assert.Throws<InvalidDataException>(() => VisioDocument.LoadLegacyBinary(source));
        Assert.True(source.CanRead);
    }

    [Fact]
    public void ReaderUsesTheNativeOwnerForPathsAndContentDetectedStreams() {
        var reader = new OfficeIMO.Reader.OfficeDocumentReaderBuilder().AddVisioHandler().Build();
        var document = reader.ReadDocument(Fixture("vsd"));
        Assert.Equal(OfficeIMO.Reader.ReaderInputKind.Visio, document.Kind);
        Assert.Contains(document.Diagnostics, diagnostic => diagnostic.Code == "VSD_CACHED_RECONSTRUCTION");
        Assert.Contains(document.Chunks, chunk => (chunk.Text ?? "").Contains("Office"));
        using var input = File.OpenRead(Fixture("vsd"));
        input.Position = 17;
        var detected = reader.ReadDocument(input, "drawing.bin", new OfficeIMO.Reader.ReaderOptions { DetectionMode = OfficeIMO.Reader.ReaderDetectionMode.PreferContent });
        Assert.Equal(OfficeIMO.Reader.ReaderInputKind.Visio, detected.Kind);
        Assert.Equal(document.Chunks.Select(chunk => chunk.Text), detected.Chunks.Select(chunk => chunk.Text));
        Assert.True(input.CanRead); Assert.Equal(17, input.Position);
        Assert.Equal(OfficeIMO.Reader.ReaderFormatSupport.ReadConvert, reader.GetCapabilities()[0].FormatQualifications.Single(profile => profile.Extension == ".vsd").Support);
    }
}
