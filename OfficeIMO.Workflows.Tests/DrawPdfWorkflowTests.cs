using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odg.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class DrawPdfWorkflowTests {
    [Fact]
    public async Task LosslessAcceptanceIncludesPdfSaveTimeLossesAndPreservesDestination() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-draw-pdf-loss-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string input = Path.Combine(directory, "source.fodg"), output = Path.Combine(directory, "result.pdf");
            var drawing = OdgDocument.Create();
            const string family = "Workflow Arabic", text = "مرحبا";
            drawing.AddPage("Unicode").Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 12, 5), string.Empty)
                .AddParagraph(text).FontFamily = family;
            drawing.SaveFlatXml(input);
            byte[] font = OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFont(text.Select(c => (int)c).Distinct().ToArray());
            var options = new OdgToPdfOptions { PdfOptions = new PdfOptions()
                .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(family, font)) };
            PdfDocumentConversionResult baseline = drawing.ToPdfDocumentResult(options);
            baseline.ToBytes();
            Assert.True(baseline.Report.HasLoss);
            byte[] original = { 1, 2, 3 };
            File.WriteAllBytes(output, original);
            var result = await new OfficeWorkflowRunner().RunAsync(new() {
                Operation = OfficeWorkflowOperation.Convert, InputPath = input, OutputPath = output,
                ConversionRouteId = "odg-pdf",
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                ConversionOptions = new() { RequireNoLoss = true, Draw = options }
            });
            Assert.False(result.Succeeded);
            var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
            foreach (OfficeConversionFidelityDiagnostic finding in baseline.Report.FidelityDiagnostics.Where(d => d.LossKind != OfficeConversionLossKind.None))
                Assert.Contains(evidence.FidelityDiagnostics, d => d.Code == finding.Code && d.Source == finding.Source && d.LossKind == finding.LossKind);
            Assert.Equal(original, File.ReadAllBytes(output));
            Assert.Empty(Directory.GetFiles(directory, ".*.tmp"));
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(".odg")]
    [InlineData(".fodg")]
    public async Task DrawRoutesPublishAllPagesAndKeepSourceAndPdfEvidence(string extension) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-draw-workflow-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string input = Path.Combine(directory, "source" + extension), output = Path.Combine(directory, "result.pdf");
            var drawing = OdgDocument.Create();
            drawing.AddPage("First", OdfLength.Points(360), OdfLength.Points(220))
                .Shapes.AddTextBox(new OdfRect(OdfLength.Points(20), OdfLength.Points(20), OdfLength.Points(200), OdfLength.Points(60)), "WorkflowCaption");
            byte[] pixels = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.Blue));
            var caption = drawing.Pages[0].Shapes.AddImage(pixels, "caption.png",
                new OdfRect(OdfLength.Points(240), OdfLength.Points(30), OdfLength.Points(70), OdfLength.Points(120))).AddParagraph("ImageCaption");
            caption.FontFamily = "Arial"; caption.FontSize = OdfLength.Points(10);
            drawing.AddPage("Empty", OdfLength.Points(180), OdfLength.Points(120));
            if (extension == ".fodg") drawing.SaveFlatXml(input); else drawing.Save(input);
            byte[] before = File.ReadAllBytes(input);
            var result = await OfficeWorkflow.Convert(input).To(output).RunAsync();
            Assert.True(result.Succeeded, result.Summary);
            Assert.Equal("odg-pdf", OfficeWorkflowCatalog.Find(extension, ".pdf")?.Id);
            var read = PdfReadDocument.Open(File.ReadAllBytes(output));
            Assert.Equal(2, read.Pages.Count);
            Assert.Equal((180D, 120D), read.Pages[1].GetPageSize());
            Assert.Contains("WorkflowCaption", read.ExtractText());
            Assert.Contains("ImageCaption", read.ExtractText());
            OfficeWorkflowConversionEvidence evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
            Assert.Equal("2", evidence.Facts["sourcePages"]);
            Assert.Equal(extension == ".fodg" ? "FODG" : "ODG", evidence.Facts["sourceFormat"]);
            Assert.Contains(evidence.FidelityDiagnostics, d => d.Location == "page:2:Empty");
            Assert.True(evidence.HasLoss);
            Assert.Contains(result.Diagnostics, d => d.Code == "OutputReopened");
            Assert.Equal(before, File.ReadAllBytes(input));
            var preview = OfficeWorkflowRunner.PreviewDocument(before, extension);
            Assert.Equal(2, preview.Pages.Count);
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void DrawSettingsSnapshotAndMixedBatchSelectionStayRouteSpecific() {
        var settings = new OfficeWorkflowConversionOptions { Draw = new OdgToPdfOptions { ForPrint = false, MaximumPages = 3 } };
        var copy = settings.ForRoute("odg-pdf");
        settings.Draw.MaximumPages = 50;
        Assert.Equal(3, copy.Draw!.MaximumPages);
        Assert.Null(settings.ForRoute("txt-pdf").Draw);
    }

    [Theory]
    [InlineData(".odg")]
    [InlineData(".fodg")]
    public async Task StrictDrawRejectionRetainsLossEvidenceAndPreservesDestination(string extension) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-draw-rejection-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string input = Path.Combine(directory, "source" + extension), output = Path.Combine(directory, "result.pdf");
            var drawing = OdgDocument.Create();
            drawing.Metadata.Title = "Unprojected metadata";
            drawing.AddPage("Rejected");
            if (extension == ".fodg") drawing.SaveFlatXml(input); else drawing.Save(input);
            byte[] original = { 1, 2, 3 };
            File.WriteAllBytes(output, original);
            var result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert, InputPath = input, OutputPath = output, ConversionRouteId = "odg-pdf",
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                ConversionOptions = new OfficeWorkflowConversionOptions { Draw = new() { LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported } }
            });
            Assert.False(result.Succeeded);
            var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
            Assert.True(evidence.HasLoss);
            Assert.Equal("1", evidence.Facts["sourcePages"]);
            Assert.Contains(evidence.FidelityDiagnostics, d => d.Location == "document-metadata");
            Assert.Contains(result.Diagnostics, d => d.Code == "ODF_SKIPPED" && d.Details!["location"] == "document-metadata");
            Assert.Equal(original, File.ReadAllBytes(output));
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PageAndSerializedOutputLimitsPreventPublication(bool pageLimit) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-draw-limit-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string input = Path.Combine(directory, "source.fodg"), output = Path.Combine(directory, "result.pdf");
            var drawing = OdgDocument.Create(); drawing.AddPage("One"); drawing.AddPage("Two"); drawing.SaveFlatXml(input);
            var result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert, InputPath = input, OutputPath = output, ConversionRouteId = "odg-pdf",
                Limits = new OfficeWorkflowLimits { MaximumOutputBytes = pageLimit ? 1024 * 1024 : 64 },
                ConversionOptions = new OfficeWorkflowConversionOptions { Draw = new() { MaximumPages = pageLimit ? 1 : 2 } }
            });
            Assert.False(result.Succeeded);
            Assert.False(File.Exists(output));
            if (!pageLimit) {
                Assert.Equal(OfficeWorkflowFailureKind.OutputFailed, result.FailureKind);
                Assert.Equal("2", Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence).Facts["sourcePages"]);
            }
        } finally { Directory.Delete(directory, true); }
    }
}
