using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioBackgroundCompositionTests {
    [Fact]
    public void ChainedBackgroundsPaintBehindForegroundWithoutResizingOrChangingSource() {
        VisioDocument source = CreateChain();
        foreach (VisioDocument document in Reopened(source)) {
            byte[] before = document.ToLegacyXmlResult().Value;
            VisioPage foreground = document.Pages[2];
            OfficeRasterImage native = Decode(foreground.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = 72, Supersampling = 1, RenderText = false, RenderStencilArtwork = false
            }));
            OfficeDrawing scene = foreground.ToDrawing().Value;
            OfficeRasterImage drawing = Decode(OfficeDrawingRasterRenderer.ToPng(scene, background: OfficeColor.White));
            Assert.Equal((288, 216), (native.Width, native.Height));
            foreach (OfficeRasterImage image in new[] { native, drawing }) {
                Assert.Equal(OfficeColor.Red, image.GetPixel(7, 209));
                Assert.Equal(OfficeColor.Blue, image.GetPixel(72, 144));
                Assert.Equal(OfficeColor.Green, image.GetPixel(108, 108));
                // A background's off-page artwork can still fall inside the foreground surface.
                Assert.Equal(OfficeColor.Red, image.GetPixel(252, 180));
            }
            string svg = foreground.ToSvg(new VisioSvgSaveOptions { RenderText = false });
            Assert.Contains("data-officeimo-visio-background=\"Base\"", svg);
            Assert.Contains("data-officeimo-visio-background=\"Overlay\"", svg);
            Assert.True(svg.IndexOf("Base", StringComparison.Ordinal) < svg.IndexOf("Overlay", StringComparison.Ordinal));
            Assert.Equal(3, document.ToDrawings().Value.Count);
            var pdf = PdfReadDocument.Open(document.ToPdfBytes(new VisioToPdfOptions { Mode = VisioPdfProjectionMode.DiagramPages }));
            Assert.Equal(3, pdf.Pages.Count);
            Assert.Equal((288D, 216D), pdf.Pages[2].GetPageSize());
            Assert.Contains("Base label", pdf.Pages[2].ExtractText());
            Assert.Contains("Overlay label", pdf.Pages[2].ExtractText());
            Assert.Equal(before, document.ToLegacyXmlResult().Value);
        }
    }

    [Theory]
    [InlineData(VisioLayerRenderMode.Visible, "Screen background", "Print background")]
    [InlineData(VisioLayerRenderMode.Printable, "Print background", "Screen background")]
    public void BackgroundLayersAreResolvedOnTheirOwnPage(VisioLayerRenderMode mode, string included, string excluded) {
        VisioDocument document = VisioDocument.Create(); document.UseMastersByDefault = false;
        VisioPage background = document.AddBackgroundPage("Background", 4, 3);
        background.AddLayer("Screen").Print = false;
        background.AddLayer("Print").Visible = false;
        background.AddRectangle(1, 1, 1.5, 1, "Screen background").LayerNames.Add("Screen");
        background.AddRectangle(3, 1, 1.5, 1, "Print background").LayerNames.Add("Print");
        VisioPage foreground = document.AddPage("Foreground", 4, 3).SetBackgroundPage(background);
        foreground.AddLayer("Screen").Visible = false;
        foreground.AddLayer("Print").Print = false;
        string svg = foreground.ToSvg(new VisioSvgSaveOptions { LayerMode = mode });
        Assert.Contains(included, svg); Assert.DoesNotContain(excluded, svg);
        var result = document.ToPdfDocumentResult(new VisioToPdfOptions {
            Mode = VisioPdfProjectionMode.DiagramPages, DrawingOptions = new VisioDrawingOptions { LayerMode = mode }
        });
        string text = PdfReadDocument.Open(result.ToBytes()).Pages[1].ExtractText();
        Assert.Contains(included, text); Assert.DoesNotContain(excluded, text);
    }

    [Fact]
    public void RepeatedBackgroundInstancesConsumeOperationBudgetsAndSinglePageChainLimit() {
        VisioDocument document = CreateChain();
        Assert.Throws<InvalidDataException>(() => document.Pages[2].ToDrawing(new VisioDrawingOptions { MaximumPages = 2 }));
        // Base: 3 shapes; Overlay: Base + 2; Foreground: Base + Overlay + 1 = 14 instances.
        Assert.Throws<InvalidDataException>(() => document.ToDrawings(new VisioDrawingOptions { MaximumShapes = 13 }));
        Assert.Equal(3, document.ToDrawings(new VisioDrawingOptions { MaximumShapes = 14 }).Value.Count);
        Assert.Throws<OperationCanceledException>(() => document.Pages[2].ToDrawing(cancellationToken: new CancellationToken(true)));
    }

    [Fact]
    public void SetterRejectsIndirectCyclesBeforeMutatingEitherPage() {
        VisioDocument document = CreateChain();
        VisioPage first = document.Pages[0], last = document.Pages[2];
        Assert.Throws<InvalidOperationException>(() => first.SetBackgroundPage(last));
        Assert.Null(first.BackgroundPage); Assert.False(last.IsBackground);
    }

    [Fact]
    public void ExistingVisioProducedBackgroundsReachForegroundSvgAndSearchablePdf() {
        string fixture = Path.Combine(RepositoryTestPaths.Find(), "Assets", "VisioTemplates", "DrawingWithLotsOfShapresAndArrows.vsdx");
        VisioDocument document = VisioDocument.Load(fixture);
        byte[] before = document.ToLegacyXmlResult().Value;
        var pdf = PdfReadDocument.Open(document.ToPdfBytes(new VisioToPdfOptions { Mode = VisioPdfProjectionMode.DiagramPages }));
        for (int index = 0; index < document.Pages.Count; index++) {
            VisioPage page = document.Pages[index];
            if (page.BackgroundPage == null) continue;
            VisioPage background = page.BackgroundPage!;
            XDocument svg = XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions { RenderText = false }));
            XElement group = Assert.Single(svg.Descendants(), element => (string?)element.Attribute("data-officeimo-visio-background") == background.Name);
            Assert.Contains(group.Descendants(), element => element.Attribute("data-visio-shape-id") != null);
            foreach (VisioShape shape in background.Shapes.Where(shape => !string.IsNullOrWhiteSpace(shape.Text))) {
                string firstWord = shape.Text!.Split(new[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)[0];
                Assert.Contains(firstWord, pdf.Pages[index].ExtractText());
            }
        }
        Assert.Equal(before, document.ToLegacyXmlResult().Value);
    }

    [Theory]
    [InlineData("2", "VISIO_BACKGROUND_CYCLE")]
    [InlineData("0", "VISIO_BACKGROUND_CYCLE")]
    [InlineData("99999", "VISIO_BACKGROUND_MISSING")]
    public void MalformedLoadedReferencesStayPreservedAndExposeLossAcrossProjectionOwners(string reference, string code) {
        XDocument xml = XDocument.Load(new MemoryStream(CreateChain().ToLegacyXmlResult().Value));
        XElement first = xml.Descendants().First(element => element.Name.LocalName == "Page");
        first.SetAttributeValue("BackPage", reference);
        VisioDocument document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml.ToString()))).Value;
        byte[] before = document.ToLegacyXmlResult().Value;
        var result = document.Pages[2].ToDrawing();
        Assert.Contains(result.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == code && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Throws<OfficeConversionException>(() => document.Pages[2].ToDrawing(new VisioDrawingOptions { RequireNoLoss = true }));
        foreach (OfficeImageExportFormat format in new[] { OfficeImageExportFormat.Svg, OfficeImageExportFormat.Png }) {
            var exported = document.Pages[2].ExportImage(format, new VisioImageExportOptions { RenderText = false, TargetDpi = 24, Supersampling = 1 });
            Assert.Contains(exported.Diagnostics, diagnostic => diagnostic.Code == code && diagnostic.LossKind == OfficeConversionLossKind.Omission);
            Assert.Throws<OfficeImageExportPolicyException>(() => document.Pages[2].ExportImage(format, new VisioImageExportOptions {
                RenderText = false, TargetDpi = 24, Supersampling = 1, Policy = new OfficeImageExportPolicy { RequireNoLoss = true }
            }));
        }
        Assert.Equal(before, document.ToLegacyXmlResult().Value);
    }

    private static VisioDocument CreateChain() {
        VisioDocument document = VisioDocument.Create(); document.UseMastersByDefault = false;
        VisioPage background = document.AddBackgroundPage("Base", 4, 4);
        background.PageScale = new VisioScaleSetting(.5, VisioMeasurementUnit.Inches);
        background.DrawingScale = new VisioScaleSetting(1, VisioMeasurementUnit.Inches);
        AddColor(background, 2, 2, 4, 4, OfficeColor.Red);
        AddColor(background, 7, 1, .5, .5, OfficeColor.Red);
        background.AddRectangle(1.5, 3.5, 2, .8, "Base label");
        VisioPage overlay = document.AddBackgroundPage("Overlay", 3, 2).SetBackgroundPage(background);
        AddColor(overlay, 1, 1, 1.5, 1.5, OfficeColor.Blue);
        overlay.AddRectangle(2, 1.75, 1.7, .4, "Overlay label");
        VisioPage foreground = document.AddPage("Foreground", 4, 3).SetBackgroundPage(overlay);
        AddColor(foreground, 1.5, 1.5, .5, .5, OfficeColor.Green);
        return document;
    }

    private static void AddColor(VisioPage page, double x, double y, double width, double height, OfficeColor color) {
        VisioShape shape = page.AddRectangle(x, y, width, height, string.Empty);
        shape.FillColor = color; shape.LinePattern = 0;
    }

    private static IEnumerable<VisioDocument> Reopened(VisioDocument document) {
        yield return document;
        yield return VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value;
        yield return VisioDocument.Load(new MemoryStream(document.ToBytes()));
    }

    private static OfficeRasterImage Decode(byte[] bytes) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out OfficeRasterImage? image)); return image!;
    }
}
