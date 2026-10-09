using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Visio.Pdf;
using Xunit;

namespace OfficeIMO.Visio.Tests;

public sealed class VisioDrawingProjectionTests {
    [Theory]
    [InlineData(VisioPackageType.Drawing)]
    [InlineData(VisioPackageType.Template)]
    public void PhysicalPageSizesBlankPagesAndCachedScaleSurviveBothSerializationFamilies(VisioPackageType family) {
        VisioDocument source = VisioDocument.Create(family);
        VisioPage first = source.AddPage("Scaled", 12, 8);
        first.PageScale = new VisioScaleSetting(0.25, VisioMeasurementUnit.Inches);
        first.DrawingScale = new VisioScaleSetting(1, VisioMeasurementUnit.Inches);
        VisioShape shape = first.AddRectangle(6, 4, 4, 2, "Searchable label");
        shape.TextStyle = new VisioTextStyle { FontFamily = "Arial", Size = 10 };
        source.AddPage("Tiny", 0.001, 0.002);
        source.AddPage("Blank", 2, 3);
        foreach (VisioDocument candidate in Reopened(source)) {
            byte[] before = candidate.ToLegacyXmlResult().Value;
            var projected = candidate.ToDrawings();
            Assert.Equal(3, projected.Value.Count);
            Assert.Equal((216D, 144D), (projected.Value[0].Width, projected.Value[0].Height));
            Assert.Equal(0.072, projected.Value[1].Width, 9);
            Assert.Equal(0.144, projected.Value[1].Height, 9);
            Assert.Empty(projected.Value[1].Elements);
            PdfDocumentConversionResult conversion = candidate.ToPdfDocumentResult(DiagramOptions());
            PdfReadDocument pdf = PdfReadDocument.Open(conversion.ToBytes());
            Assert.Equal(3, pdf.Pages.Count);
            Assert.Equal((216D, 144D), pdf.Pages[0].GetPageSize());
            Assert.Equal((144D, 216D), pdf.Pages[2].GetPageSize());
            Assert.Contains("Searchable label", NormalizeText(pdf.ExtractText()));
            Assert.DoesNotContain(conversion.Warnings, warning => warning.Code == "pdf-projection-visio-semantic-fallback");
            Assert.IsType<VisioDrawingConversionReport>(Assert.Single(conversion.SourceConversionReports));
            Assert.Equal(before, candidate.ToLegacyXmlResult().Value);
        }
    }

    [Fact]
    public void ScreenAndPrintSelectionKeepsChildrenAndConnectorsIndependentOfTheirParentAndEndpoints() {
        VisioDocument document = VisioDocument.Create();
        VisioPage page = document.AddPage("Layers", 5, 4);
        VisioLayer hidden = page.AddLayer("Hidden"); hidden.Visible = false; hidden.Print = false;
        VisioLayer print = page.AddLayer("Print"); print.Visible = false; print.Print = true;
        var group = new VisioShape("group") { Type = "Group", PinX = 0, PinY = 0, LocPinX = 0, LocPinY = 0 };
        group.LayerNames.Add(hidden.Name);
        var child = new VisioShape("child", 1, 1, 1, 1, "Visible child");
        group.Children.Add(child); page.Shapes.Add(group);
        VisioShape printOnly = page.AddRectangle(3, 2, 1, 1, "Print only"); printOnly.LayerNames.Add(print.Name);
        VisioConnector connector = page.AddConnector("route", child, printOnly, ConnectorKind.Straight);
        connector.Label = "Independent route";
        string PrintText(VisioLayerRenderMode mode) => PdfReadDocument.Open(document.ToPdfBytes(new VisioToPdfOptions {
            Mode = VisioPdfProjectionMode.DiagramPages, DrawingOptions = new VisioDrawingOptions { LayerMode = mode }
        })).ExtractText();
        Assert.Contains("Visible child", PrintText(VisioLayerRenderMode.Visible));
        Assert.DoesNotContain("Print only", PrintText(VisioLayerRenderMode.Visible));
        Assert.Contains("Independent route", PrintText(VisioLayerRenderMode.Visible));
        Assert.Contains("Print only", PrintText(VisioLayerRenderMode.Printable));
    }

    [Fact]
    public void ObjectAndPointBudgetsAreOperationWideAndIncludeExcludedObjects() {
        VisioDocument document = VisioDocument.Create();
        VisioPage first = document.AddPage("First");
        VisioLayer hidden = first.AddLayer("Hidden"); hidden.Print = false;
        first.AddRectangle(1, 1, 1, 1, "Hidden").LayerNames.Add(hidden.Name);
        document.AddPage("Second").AddRectangle(1, 1, 1, 1, "Second");
        Assert.Throws<InvalidDataException>(() => document.ToDrawings(new VisioDrawingOptions { MaximumPages = 1 }));
        Assert.Throws<InvalidDataException>(() => document.ToDrawings(new VisioDrawingOptions { MaximumShapes = 1 }));
        Assert.Throws<InvalidDataException>(() => document.ToDrawings(new VisioDrawingOptions { MaximumGeometryPoints = 1 }));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => document.ToDrawings(cancellationToken: cancellation.Token));
    }

    [Fact]
    public void StencilProjectionRequiresExplicitInstantiationAndModeSettingsCannotBeSilentlyIgnored() {
        VisioDocument stencil = VisioDocument.Create(VisioPackageType.Stencil);
        stencil.RegisterMaster("Box", new VisioShape("master", 1, 1, 1, 1, "Reusable"));
        Assert.Throws<InvalidOperationException>(() => stencil.ToDrawings());
        Assert.Throws<InvalidOperationException>(() => stencil.ToPdfBytes(DiagramOptions()));
        VisioDocument drawing = VisioDocument.Create(); drawing.AddPage("Page");
        Assert.Throws<ArgumentException>(() => drawing.ToPdfDocumentResult(new VisioToPdfOptions { DrawingOptions = new VisioDrawingOptions() }));
        Assert.Throws<ArgumentException>(() => drawing.ToPdfDocumentResult(new VisioToPdfOptions { Mode = VisioPdfProjectionMode.DiagramPages, ProjectionOptions = new PdfProjectionOptions() }));
        Assert.Throws<ArgumentOutOfRangeException>(() => drawing.ToPdfDocumentResult(new VisioToPdfOptions { Mode = (VisioPdfProjectionMode)100 }));
    }

    [Fact]
    public void LocatedSourceLossSurvivesPdfResultsAndStrictSavingLeavesDestinationsUntouched() {
        VisioDocument document = VisioDocument.Create();
        document.AddPage("Page").AddRectangle(1, 1, 1, 1, "Label").AddHyperlink("https://example.com");
        var result = document.ToPdfDocumentResult(new VisioToPdfOptions { Mode = VisioPdfProjectionMode.DiagramPages, SourceName = "source.vdx" });
        VisioDrawingConversionReport report = Assert.IsType<VisioDrawingConversionReport>(Assert.Single(result.SourceConversionReports));
        Assert.Equal("source.vdx", report.SourceName);
        Assert.Contains(result.FidelityDiagnostics, diagnostic => diagnostic.Code == "VISIO_DRAWING_METADATA" && diagnostic.LossKind == OfficeConversionLossKind.Omission && diagnostic.Location!.StartsWith("page:1:Page:shape:"));
        byte[] sentinel = { 1, 2, 3 }; using var destination = new MemoryStream(); destination.Write(sentinel, 0, sentinel.Length);
        Assert.Throws<OfficeConversionException>(() => result.SaveLossless(destination));
        Assert.Equal(sentinel, destination.ToArray());
        Assert.Throws<OfficeConversionException>(() => document.ToPdfDocumentResult(new VisioToPdfOptions {
            Mode = VisioPdfProjectionMode.DiagramPages, DrawingOptions = new VisioDrawingOptions { RequireNoLoss = true }
        }));
    }

    [Theory]
    [InlineData("clickhouse-distributed-insert.vdx", "INSERT INTO")]
    [InlineData("nxbre-chocolatebox.vdx", "binder")]
    public void IndependentProducerRunsProjectToSearchablePageContentWithoutChangingNativeXml(string name, string text) {
        var document = VisioDocument.LoadLegacyXml(Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", name)).Value;
        foreach (VisioDocument candidate in Reopened(document)) {
            byte[] before = candidate.ToLegacyXmlResult().Value;
            VisioToPdfOptions options = DiagramOptions();
            options.DrawingOptions = new VisioDrawingOptions();
            byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf"));
            foreach (string family in new[] { "Yandex Sans Text", "Monaco", "Arial" }) options.DrawingOptions.Fonts.Add(family, font);
            PdfDocumentConversionResult conversion = candidate.ToPdfDocumentResult(options);
            PdfReadDocument pdf = PdfReadDocument.Open(conversion.ToBytes());
            Assert.Equal(candidate.Pages.Count, pdf.Pages.Count);
            Assert.Contains(text, NormalizeText(pdf.ExtractText()), StringComparison.OrdinalIgnoreCase);
            Assert.Contains(conversion.FidelityDiagnostics, diagnostic => diagnostic.Code == "VISIO_DRAWING_TEXT_LAYOUT");
            Assert.Equal(before, candidate.ToLegacyXmlResult().Value);
        }
    }

    [Fact]
    public void AngularBitmapStaysAnImageInDiagramPdfAndRetainsDecodeDiagnostics() {
        var document = VisioDocument.LoadLegacyXml(Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", "angularjs-simple-scope.vdx")).Value;
        byte[] before = document.ToLegacyXmlResult().Value;
        var conversion = document.ToPdfDocumentResult(DiagramOptions());
        PdfExtractedImage image = Assert.Single(PdfReadDocument.Open(conversion.ToBytes()).Pages.SelectMany(page => page.GetImages()));
        Assert.True(OfficeRasterImageDecoder.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.Equal((113, 155), (raster!.Width, raster.Height));
        Assert.Contains(conversion.FidelityDiagnostics, diagnostic => diagnostic.Code == "VISIO_BITMAP_SIZE_NORMALIZED" && diagnostic.LossKind == OfficeConversionLossKind.None);
        Assert.Throws<InvalidDataException>(() => document.ToDrawings(new VisioDrawingOptions { MaximumTotalImageBytes = 1 }));
        Assert.Equal(before, document.ToLegacyXmlResult().Value);
        VisioDocument repeated = VisioDocument.LoadLegacyXml(new MemoryStream(before)).Value;
        repeated.DuplicatePage(repeated.Pages[0], "Repeated image");
        Assert.Throws<InvalidDataException>(() => repeated.ToDrawings(new VisioDrawingOptions { MaximumTotalImagePixels = 113 * 155 }));
    }

    [Fact]
    public void JustifiedParagraphsReachTheSharedParagraphLayout() {
        VisioPage page = VisioDocument.Create().AddPage("Justified");
        VisioShape shape = page.AddRectangle(2, 2, 3, 2, "One two three\nfour five six");
        shape.TextStyle = new VisioTextStyle { HorizontalAlignment = VisioTextHorizontalAlignment.Justify };
        OfficeDrawing scene = page.ToDrawing().Value;
        OfficeDrawingEffectGroup group = Assert.Single(scene.Elements.OfType<OfficeDrawingEffectGroup>());
        OfficeDrawingRichText text = Assert.Single(group.Drawing.Elements.OfType<OfficeDrawingRichText>());
        Assert.Equal(OfficeTextAlignment.Justify, Assert.Single(text.Paragraphs!).Alignment);
    }

    [Fact]
    public void FittedTextRemainsSearchableAndImpossibleFramesReportLocatedOmission() {
        VisioDocument source = VisioDocument.Create(); source.UseMastersByDefault = false;
        VisioPage page = source.AddPage("Fitting");
        page.AddRectangle(2, 2, 2, .3, "Fit me fully").TextStyle = new VisioTextStyle { Size = 24 };
        page.AddRectangle(4, 4, 1, .01, "Cannot fit").TextStyle = new VisioTextStyle {
            Size = 24, TopMargin = 0, BottomMargin = 0
        };
        foreach (VisioDocument document in Reopened(source)) {
            byte[] before = document.ToLegacyXmlResult().Value;
            var result = document.ToPdfDocumentResult(DiagramOptions());
            Assert.Contains("Fit me fully", NormalizeText(PdfReadDocument.Open(result.ToBytes()).ExtractText()));
            Assert.Contains(result.FidelityDiagnostics, item => item.Code == "VISIO_DRAWING_TEXT_CLIPPED" &&
                item.LossKind == OfficeConversionLossKind.Omission && item.Location!.StartsWith("page:1:Fitting:shape:"));
            Assert.Equal(before, document.ToLegacyXmlResult().Value);
        }
    }

    private static VisioToPdfOptions DiagramOptions() => new() { Mode = VisioPdfProjectionMode.DiagramPages };

    private static string NormalizeText(string text) => string.Join(" ", text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    private static IEnumerable<VisioDocument> Reopened(VisioDocument source) {
        yield return source;
        yield return VisioDocument.LoadLegacyXml(new MemoryStream(source.ToLegacyXmlResult().Value)).Value;
        yield return VisioDocument.Load(new MemoryStream(source.ToBytes()));
    }
}
