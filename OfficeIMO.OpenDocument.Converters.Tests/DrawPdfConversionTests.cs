using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odg.Pdf;
using OfficeIMO.OpenDocument.Testing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class DrawPdfConversionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PackageAndFlatRoutesPreserveMixedPageSizesOrderAndBlankPages(bool flat) {
        OdgDocument source = CreateDrawing();
        using var input = new MemoryStream();
        if (flat) source.SaveFlatXml(input); else source.Save(input);
        input.Position = 0;
        OdgDocument loaded = flat ? OdgDocument.LoadFlatXml(input) : OdgDocument.Load(input);
        var conversion = loaded.ToPdfDocumentResult();
        byte[] bytes = conversion.ToBytes();
        PdfReadDocument read = PdfReadDocument.Open(bytes);
        Assert.Equal(3, read.Pages.Count);
        Assert.Equal((360D, 220D), read.Pages[0].GetPageSize());
        Assert.Equal((220D, 360D), read.Pages[1].GetPageSize());
        Assert.Equal((180D, 120D), read.Pages[2].GetPageSize());
        Assert.Contains("FirstCaption", read.Pages[0].ExtractText());
        Assert.Contains("SecondCaption", read.Pages[1].ExtractText());
        Assert.Equal(string.Empty, read.Pages[2].ExtractText());
        OdfConversionReport report = Assert.IsType<OdfConversionReport>(Assert.Single(conversion.SourceConversionReports));
        Assert.Contains(report.Mappings, m => m.Feature == "page:3:Blank" && m.Status == OdfConversionMappingStatus.Converted);
        Assert.True(input.CanRead);
    }

    [Fact]
    public void LayerIntentUsesPrintByDefaultAndScreenWhenRequested() {
        OdgDocument source = OdgDocument.Create();
        source.Layers.Add("screen", OdgLayerDisplay.Screen);
        source.Layers.Add("print", OdgLayerDisplay.Printer);
        OdgPage page = source.AddPage();
        page.Shapes.AddTextBox(new OdfRect(OdfLength.Points(20), OdfLength.Points(20), OdfLength.Points(150), OdfLength.Points(50)), "ScreenOnly").Layer = "screen";
        page.Shapes.AddTextBox(new OdfRect(OdfLength.Points(20), OdfLength.Points(80), OdfLength.Points(150), OdfLength.Points(50)), "PrintOnly").Layer = "print";
        string printed = PdfReadDocument.Open(source.ToPdfBytes()).ExtractText();
        string screened = PdfReadDocument.Open(source.ToPdfBytes(new() { ForPrint = false })).ExtractText();
        Assert.Contains("PrintOnly", printed); Assert.DoesNotContain("ScreenOnly", printed);
        Assert.Contains("ScreenOnly", screened); Assert.DoesNotContain("PrintOnly", screened);
    }

    [Fact]
    public void UnsupportedContentHasPageLocationsAndStrictPolicyLeavesOutputUntouched() {
        OdgDocument source = CreateDrawing();
        source.Pages[1].Shapes.AddRectangle(OdfRect.FromCentimeters(1, 4, 2, 2), "UnsupportedGeometry").Element.Name = OdfNamespaces.Draw + "custom-shape";
        var conversion = source.ToPdfDocumentResult();
        Assert.Contains(conversion.FidelityDiagnostics, d => d.Location!.StartsWith("page:2:Second/") && d.LossKind == OfficeConversionLossKind.Omission);
        Assert.Throws<OdfConversionLossException>(() => conversion.RequireNoLoss());
        using var output = new MemoryStream(new byte[] { 1, 2, 3 }, writable: true);
        var result = source.SaveAsPdfResult(output, new() { LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported });
        Assert.False(result.Succeeded);
        Assert.Equal(new byte[] { 1, 2, 3 }, output.ToArray());
    }

    [Fact]
    public void EmptyPageLimitInvalidSettingsAndCancellationFailBeforeSaving() {
        Assert.Throws<InvalidOperationException>(() => OdgDocument.Create().ToPdfDocumentResult());
        OdgDocument source = CreateDrawing();
        Assert.Throws<InvalidOperationException>(() => source.ToPdfDocumentResult(new() { MaximumPages = 2 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => source.ToPdfDocumentResult(new() { MaximumPages = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => source.ToPdfDocumentResult(new() { LossPolicy = (OdfConversionLossPolicy)42 }));
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => source.ToPdfDocumentResult(cancellationToken: canceled.Token));
        Assert.Throws<OperationCanceledException>(() => source.Pages[0].ToDrawing(OdfConversionLossPolicy.ReportOnly, true, canceled.Token));
    }

    [Fact]
    public async Task SaveLifecyclePreservesSourceAndRefreshesPdfStageWarnings() {
        OdgDocument source = CreateDrawing();
        string xml = source.GetXml("content.xml").ToString(SaveOptions.DisableFormatting);
        var settings = new OdgToPdfOptions { PdfOptions = new PdfOptions { PageSize = PageSizes.Letter } };
        var conversion = source.ToPdfDocumentResult(settings);
        // A caller may add PDF-native content before saving; its save-time diagnostics must stay linked.
        var vertical = new OfficeDrawing(100, 100).AddVerticalText("Vertical", 0, 0, 100, 100);
        conversion.Value.Compose(builder => builder.Page(page => page.Canvas(canvas => canvas.Drawing(vertical, 0, 0, 100, 100))));
        using var output = new MemoryStream();
        PdfSaveResult saved = await conversion.SaveAsync(output);
        saved.RequireSuccess();
        Assert.Equal(2, saved.ConversionReports.Count);
        Assert.True(saved.HasLoss);
        Assert.Contains(conversion.FidelityDiagnostics, d => d.Source.StartsWith("OfficeIMO.OpenDocument:"));
        Assert.Contains(conversion.Warnings, w => w.Code == "vertical-text-stacked-fallback");
        Assert.Equal(xml, source.GetXml("content.xml").ToString(SaveOptions.DisableFormatting));
        Assert.Equal(PageSizes.Letter, settings.PdfOptions.PageSize);
        Assert.True(output.CanWrite);
    }

    [Fact]
    public void ImageCaptionsProjectWhileDocumentMetadataRemainsAnExplicitOmission() {
        OdgDocument source = CreateDrawing();
        source.Metadata.Title = "Drawing title";
        byte[] png = OfficeDrawingRasterRenderer.ToPng(new OfficeDrawing(2, 2).AddShape(OfficeShape.Rectangle(2, 2), 0, 0));
        var image = source.Pages[0].Shapes.AddImage(png, "caption.png", OdfRect.FromCentimeters(2, 3, 6, 2), "ImageCaption");
        var paragraph = image.AddParagraph("ImageCaptionMarker");
        paragraph.FontFamily = "Arial"; paragraph.FontSize = OdfLength.Points(10);
        var projected = source.ToDrawings(forPrint: true);
        Assert.Contains(projected.Report.Mappings, m => m.Feature == "document-metadata" && m.Status == OdfConversionMappingStatus.Skipped);
        Assert.Contains(projected.Report.Mappings, m => m.Feature == "page:1:First/shape:ImageCaption:text" && m.Status == OdfConversionMappingStatus.Approximated);
        var conversion = source.ToPdfDocumentResult();
        Assert.Contains("ImageCaptionMarker", PdfReadDocument.Open(conversion.ToBytes()).ExtractText());
        Assert.DoesNotContain(conversion.FidelityDiagnostics, d => d.Location == "page:1:First/shape:ImageCaption:image-text");
        Assert.Contains(conversion.FidelityDiagnostics, d => d.Location == "document-metadata" && d.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void PackageWithoutOptionalMetadataConvertsWithoutFalseLosses() {
        OdgDocument source = CreateDrawing();
        using var input = new MemoryStream();
        source.Save(input);
        byte[] package = OdfTestPackageRewriter.Remove(input.ToArray(), "meta.xml");
        package = OdfTestPackageRewriter.Rewrite(package, (name, bytes) => {
            if (name != "META-INF/manifest.xml") return bytes;
            using var manifestInput = new MemoryStream(bytes);
            var manifest = XDocument.Load(manifestInput);
            manifest.Root!.Elements().Where(e => (string?)e.Attribute(OdfNamespaces.Manifest + "full-path") == "meta.xml").Remove();
            using var manifestOutput = new MemoryStream();
            manifest.Save(manifestOutput);
            return manifestOutput.ToArray();
        });
        using var withoutMetadata = new MemoryStream(package);
        OdgDocument loaded = OdgDocument.Load(withoutMetadata);
        Assert.True(loaded.Validate().IsValid);
        string before = loaded.GetXml("content.xml").ToString(SaveOptions.DisableFormatting);
        var conversion = loaded.ToPdfDocumentResult();
        Assert.Equal(3, PdfReadDocument.Open(conversion.ToBytes()).Pages.Count);
        Assert.Equal(before, loaded.GetXml("content.xml").ToString(SaveOptions.DisableFormatting));
        Assert.DoesNotContain(conversion.FidelityDiagnostics, d => d.Location == "document-metadata");
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task StrictSaveFailuresRetainPageQualifiedLosses(bool useAsync, bool path) {
        OdgDocument source = CreateDrawing();
        source.Pages[1].Shapes.AddRectangle(OdfRect.FromCentimeters(1, 4, 2, 2), "Rejected")
            .Element.Name = OdfNamespaces.Draw + "custom-shape";
        var options = new OdgToPdfOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported };
        byte[] original = { 1, 2, 3 };
        string outputPath = Path.Combine(Path.GetTempPath(), "officeimo-draw-strict-" + Guid.NewGuid().ToString("N") + ".pdf");
        using var output = new MemoryStream(original.ToArray(), writable: true);
        try {
            if (path) File.WriteAllBytes(outputPath, original);
            PdfSaveResult result = path
                ? useAsync ? await source.SaveAsPdfResultAsync(outputPath, options) : source.SaveAsPdfResult(outputPath, options)
                : useAsync ? await source.SaveAsPdfResultAsync(output, options) : source.SaveAsPdfResult(output, options);
            Assert.False(result.Succeeded);
            Assert.True(result.HasLoss);
            Assert.Contains(result.FidelityDiagnostics, d => d.Location!.StartsWith("page:2:Second/") && d.LossKind == OfficeConversionLossKind.Omission);
            Assert.Equal(original, path ? File.ReadAllBytes(outputPath) : output.ToArray());
            Assert.True(output.CanWrite);
        } finally { if (File.Exists(outputPath)) File.Delete(outputPath); }
    }

    private static OdgDocument CreateDrawing() {
        var source = OdgDocument.Create();
        source.AddPage("First", OdfLength.Points(360), OdfLength.Points(220))
            .Shapes.AddTextBox(new OdfRect(OdfLength.Points(20), OdfLength.Points(20), OdfLength.Points(200), OdfLength.Points(60)), "FirstCaption");
        source.AddPage("Second", OdfLength.Points(220), OdfLength.Points(360))
            .Shapes.AddTextBox(new OdfRect(OdfLength.Points(20), OdfLength.Points(20), OdfLength.Points(180), OdfLength.Points(60)), "SecondCaption");
        source.AddPage("Blank", OdfLength.Points(180), OdfLength.Points(120));
        return source;
    }
}
