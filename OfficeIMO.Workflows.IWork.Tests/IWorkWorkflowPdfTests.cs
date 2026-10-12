using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows.IWork;
using Xunit;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed partial class IWorkWorkflowTests {
    [Theory]
    [InlineData("pages", "pages-pdf", "hello pages", 1)]
    [InlineData("numbers", "numbers-pdf", "42", 1)]
    [InlineData("key", "keynote-pdf", "hello keynote", 2)]
    public async Task Registered_pdf_routes_honor_verified_shared_compression(string source, string route, string text, int pages) {
        using var files = new Files(source, "pdf");
        UseCompletePdfFixture(files, source);
        OfficeWorkflowRunner runner = IWorkWorkflow.CreateRunner();
        Assert.True(runner.ConversionRoutes.Single(item => item.Id == route).SupportsPdfCompression);
        OfficeWorkflowRequest request = files.Request(route);
        request.RegisteredConversionSettings = new IWorkWorkflowSettings();
        request.ConversionOptions = new() { CompressPdfOutput = true };
        OfficeWorkflowResult result = await runner.RunAsync(request);
        Assert.True(result.Succeeded, result.Summary + "; " + string.Join("; ", result.ConversionEvidence?.FidelityDiagnostics.Select(diagnostic => diagnostic.Code + ": " + diagnostic.Message) ?? []));
        Assert.Contains(result.Diagnostics, item => item.Code == "PdfOutputCompression");
        Assert.Contains(result.Diagnostics, item => item.Code == "OutputReopened");
        PdfReadDocument pdf = PdfReadDocument.Open(File.ReadAllBytes(files.Output));
        Assert.Equal(pages, pdf.Pages.Count);
        Assert.Contains(text, pdf.ExtractText());
        Assert.Equal("iwork-editable,pdf", result.ConversionEvidence!.Facts["conversionStages"]);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Complete_keynote_pdf_rejects_omitted_native_presenter_notes(bool compress) {
        using var files = new Files("key", "pdf");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Corpus", "simple.key"), files.Input, true);
        IWorkSourceDocument source = IWorkSourceDocument.Open(files.Input);
        Assert.Contains(source.ReadKeynote().Slides, slide => !string.IsNullOrWhiteSpace(slide.PresenterNotes));
        File.WriteAllText(files.Output, "keep existing");
        OfficeWorkflowRequest request = files.Request("keynote-pdf");
        request.RegisteredConversionSettings = new IWorkWorkflowSettings();
        request.ConversionOptions = new() { CompressPdfOutput = compress };
        request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;

        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);

        Assert.False(result.Succeeded);
        Assert.Null(result.OutputPath);
        Assert.Equal(0, result.OutputBytes);
        Assert.Equal("pdf", result.ConversionEvidence!.Facts["conversionFailureStage"]);
        Assert.Equal("True", result.ConversionEvidence.Facts["pdfReportAvailable"]);
        Assert.Contains(result.ConversionEvidence.FidelityDiagnostics, diagnostic =>
            diagnostic.Code == "speaker-notes-omitted" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "SourceSnapshot");
        Assert.Equal("keep existing", File.ReadAllText(files.Output));
        Assert.Empty(Directory.GetFiles(files.Root, "*.tmp"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Complete_pdf_rejects_known_text_clipping_and_preserves_the_destination(bool nativeReference) {
        using var files = new Files("key", "pdf");
        if (nativeReference) File.Copy(Path.Combine(AppContext.BaseDirectory, "Corpus", "colors.key"), files.Input, true);
        else {
            IWorkKeynoteDocument source = IWorkKeynoteDocument.Create();
            source.AddSlide().AddText("VISIBLE alpha beta gamma delta epsilon zeta eta theta TAIL", 10, 10, 150, 30);
            source.Save(files.Input, overwrite: true);
        }
        File.WriteAllText(files.Output, "keep existing");
        OfficeWorkflowRequest request = files.Request("keynote-pdf");
        request.RegisteredConversionSettings = new IWorkWorkflowSettings();
        request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.False(result.Succeeded);
        Assert.Contains("clipped", result.Summary);
        Assert.Contains(result.ConversionEvidence!.FidelityDiagnostics, diagnostic => diagnostic.Code == "text-box-overflow");
        Assert.Equal("0", result.ConversionEvidence.Facts["omittedSourceUnitCount"]);
        Assert.Equal("keep existing", File.ReadAllText(files.Output));
        Assert.Empty(Directory.GetFiles(files.Root, "*.tmp"));
    }

    [Fact]
    public async Task Complete_pdf_publishes_resolved_internal_slide_links_and_reopens_them() {
        using var files = new Files("key", "pdf");
        byte[] link = System.Text.Encoding.UTF8.GetBytes("#slide-2");
        byte[] tree = Join(R(2, 3), R(2, 4));
        byte[] canvas = Join([13], BitConverter.GetBytes(960f), [21], BitConverter.GetBytes(540f));
        byte[] storage = Join(B(3, System.Text.Encoding.UTF8.GetBytes("Linked title")), B(11, B(1, Join(V(1, 0), R(2, 9)))));
        WritePackage(files.Input, Join(A(1, 1, R(2, 2)), A(2, 2, Join(B(3, tree), B(4, canvas))),
            A(3, 4, R(2, 5)), A(4, 4, R(2, 6)), A(5, 5, R(5, 7)), A(6, 5, []),
            A(7, 2011, Join(B(1, B(1, B(4, link))), R(2, 8))), A(8, 2001, storage), A(9, 2032, B(2, link))));
        File.WriteAllText(files.Output, "keep existing");
        OfficeWorkflowRequest request = files.Request("keynote-pdf");
        request.RegisteredConversionSettings = new IWorkWorkflowSettings();
        request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal("0", result.ConversionEvidence!.Facts["omittedSourceUnitCount"]);
        Assert.DoesNotContain(result.ConversionEvidence.FidelityDiagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);
        PdfReadDocument pdf = PdfReadDocument.Open(File.ReadAllBytes(files.Output));
        Assert.Equal(2, pdf.Pages.Count);
        Assert.Equal(2, Assert.Single(pdf.NamedDestinations, d => d.Name == "powerpoint-slide-2").PageNumber);
        PdfLinkAnnotation[] links = pdf.Pages.SelectMany(page => page.GetLinkAnnotations()).ToArray();
        Assert.NotEmpty(links);
        Assert.All(links, link => Assert.Equal("powerpoint-slide-2", link.DestinationName));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OutputReopened");
        Assert.Empty(Directory.GetFiles(files.Root, "*.tmp"));
    }

    [Fact]
    public async Task Native_Keynote_slide_links_pass_the_complete_pdf_route_with_every_slide_and_text() {
        using var files = new Files("key", "pdf");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Corpus", "slide-links.key"), files.Input, true);
        OfficeWorkflowRequest request = files.Request("keynote-pdf");
        request.RegisteredConversionSettings = new IWorkWorkflowSettings();
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal("False", result.ConversionEvidence!.Facts["partialEditableReconstruction"]);
        Assert.Equal("0", result.ConversionEvidence.Facts["omittedSourceUnitCount"]);
        Assert.DoesNotContain(result.ConversionEvidence.FidelityDiagnostics,
            diagnostic => diagnostic.LossKind is OfficeConversionLossKind.Omission or OfficeConversionLossKind.Failure);
        PdfReadDocument pdf = PdfReadDocument.Open(File.ReadAllBytes(files.Output));
        Assert.Equal(3, pdf.Pages.Count);
        string text = pdf.ExtractText();
        foreach (string expected in new[] { "Slide links", "Go to slide two", "Frame to next slide", "Back to slide one", "Final slide" })
            Assert.Contains(expected, text);
        PdfLinkAnnotation[] links = pdf.Pages.SelectMany(page => page.GetLinkAnnotations()).ToArray();
        Assert.Equal(3, links.Length);
        Assert.Equal(new int?[] { 2, 2, 1 }, links.Select(link => pdf.NamedDestinations.Single(
            destination => destination.Name == link.DestinationName).PageNumber));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OutputReopened");
        Assert.Empty(Directory.GetFiles(files.Root, "*.tmp"));
    }

    [Fact]
    public async Task Pdf_construction_failure_retains_the_completed_source_stage() {
        using var files = new Files("pages", "pdf");
        File.WriteAllText(files.Output, "keep existing");
        var settings = new IWorkWorkflowSettings();
        settings.PagesPdfOptions.MaxImagesPerParagraph = 0;
        OfficeWorkflowRequest request = files.Request("pages-pdf");
        request.RegisteredConversionSettings = settings;
        request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.False(result.Succeeded);
        Assert.Equal("pdf", result.ConversionEvidence!.Facts["conversionFailureStage"]);
        Assert.Equal("False", result.ConversionEvidence.Facts["pdfReportAvailable"]);
        Assert.Equal("0", result.ConversionEvidence.Facts["omittedSourceUnitCount"]);
        Assert.Equal("keep existing", File.ReadAllText(files.Output));
    }

    [Theory]
    [InlineData("pages", "pages-pdf")]
    [InlineData("numbers", "numbers-pdf")]
    [InlineData("key", "keynote-pdf")]
    public async Task Cancellation_after_pdf_reopen_preserves_the_destination_and_removes_staging(string source, string route) {
        using var files = new Files(source, "pdf");
        UseCompletePdfFixture(files, source);
        File.WriteAllText(files.Output, "keep existing");
        OfficeWorkflowRequest request = files.Request(route);
        request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
        request.RegisteredConversionSettings = new IWorkWorkflowSettings();
        using var cancellation = new CancellationTokenSource();
        bool reachedPublication = false;
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request,
            new InlineProgress(update => {
                if (update.Stage != "publish") return;
                reachedPublication = true;
                cancellation.Cancel();
            }), cancellation.Token);
        Assert.True(reachedPublication);
        Assert.Equal(OfficeWorkflowStatus.Cancelled, result.Status);
        Assert.Contains(result.Diagnostics, item => item.Code == "OutputReopened");
        Assert.Equal("keep existing", File.ReadAllText(files.Output));
        Assert.Empty(Directory.GetFiles(files.Root, "*.tmp"));
    }

    [Theory]
    [InlineData("pages", "pages-pdf", "hello pages", 1)]
    [InlineData("numbers", "numbers-pdf", "42", 1)]
    [InlineData("key", "keynote-pdf", "hello keynote", 2)]
    public async Task Complete_iwork_pdf_routes_reopen_every_page_and_retain_both_stages(string source, string route, string text, int pages) {
        using var files = new Files(source, "pdf");
        UseCompletePdfFixture(files, source);
        OfficeWorkflowRequest request = files.Request(route);
        request.RegisteredConversionSettings = new IWorkWorkflowSettings();
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.True(result.Succeeded, result.Summary + "; " + string.Join("; ", result.ConversionEvidence?.FidelityDiagnostics.Select(diagnostic => diagnostic.Code + ": " + diagnostic.Message) ?? []));
        PdfReadDocument pdf = PdfReadDocument.Open(File.ReadAllBytes(files.Output));
        Assert.Equal(pages, pdf.Pages.Count);
        Assert.Contains(text, pdf.ExtractText());
        Assert.Contains(result.Diagnostics, item => item.Code == "OutputReopened");
        var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
        Assert.Equal("iwork-editable,pdf", evidence.Facts["conversionStages"]);
        Assert.Equal("False", evidence.Facts["partialEditableReconstruction"]);
        Assert.Equal("0", evidence.Facts["omittedSourceUnitCount"]);
        Assert.Contains(evidence.FidelityDiagnostics, item => item.Code == "IWORK_RECORD_FIDELITY_UNASSESSED");
        Assert.Equal(Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input))),
            Assert.Single(result.Diagnostics, item => item.Code == "SourceSnapshot").Details["sha256"], ignoreCase: true);
    }

    [Theory]
    [InlineData(IWorkConversionMode.VisualOnly, false)]
    [InlineData(IWorkConversionMode.EditableOnly, true)]
    public async Task Pdf_routes_reject_preview_and_partial_policies_before_publication(IWorkConversionMode mode, bool partial) {
        using var files = new Files("pages", "pdf");
        File.WriteAllText(files.Output, "keep existing");
        OfficeWorkflowRequest request = files.Request("pages-pdf");
        request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
        request.RegisteredConversionSettings = new IWorkWorkflowSettings {
            ConversionOptions = new() { Mode = mode, AllowPartialEditableReconstruction = partial, RequireCompleteVisualCoverage = false }
        };
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.False(result.Succeeded);
        Assert.Equal("keep existing", File.ReadAllText(files.Output));
        Assert.Empty(Directory.GetFiles(files.Root, "*.tmp"));
    }

    [Fact]
    public async Task Pdf_output_limit_retains_source_evidence_and_existing_destination() {
        using var files = new Files("pages", "pdf");
        File.WriteAllText(files.Output, "keep existing");
        OfficeWorkflowRequest request = files.Request("pages-pdf");
        request.RegisteredConversionSettings = new IWorkWorkflowSettings();
        request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
        request.Limits.MaximumOutputBytes = 32;
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.False(result.Succeeded);
        Assert.NotNull(result.ConversionEvidence);
        Assert.Equal("iwork-editable,pdf", result.ConversionEvidence.Facts["conversionStages"]);
        Assert.Equal("keep existing", File.ReadAllText(files.Output));
    }

    [Fact]
    public async Task Numbers_pdf_rejects_row_selection_and_print_area_clipping() {
        using var files = new Files("numbers", "pdf");
        var settings = new IWorkWorkflowSettings();
        settings.NumbersPdfOptions.MaxRowsPerSheet = 1;
        OfficeWorkflowRequest request = files.Request("numbers-pdf");
        request.RegisteredConversionSettings = settings;
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.False(result.Succeeded);
        Assert.Contains("every sheet, row, and column", result.Summary);
        Assert.False(File.Exists(files.Output));
    }

    [Fact]
    public async Task Keynote_pdf_rejects_content_exclusion_and_preserves_existing_destination() {
        using var files = new Files("key", "pdf");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Corpus", "simple.key"), files.Input, true);
        File.WriteAllText(files.Output, "keep existing");
        var settings = new IWorkWorkflowSettings();
        settings.KeynotePdfOptions.IncludeTextBoxes = false;
        OfficeWorkflowRequest request = files.Request("keynote-pdf");
        request.RegisteredConversionSettings = settings;
        request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.False(result.Succeeded);
        Assert.Contains("content type", result.Summary);
        Assert.Equal("keep existing", File.ReadAllText(files.Output));
    }

    [Fact]
    public async Task Pdf_font_preflight_failure_retains_stages_and_does_not_publish_incomplete_text() {
        using var files = new Files("key", "pdf");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Corpus", "unicode.key"), files.Input, true);
        File.WriteAllText(files.Output, "keep existing");
        OfficeWorkflowRequest request = files.Request("keynote-pdf");
        var settings = new IWorkWorkflowSettings();
        settings.KeynotePdfOptions.ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic();
        settings.KeynotePdfOptions.TextFallbacks = PdfTextFallbackFeatures.None;
        request.RegisteredConversionSettings = settings;
        request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.False(result.Succeeded);
        Assert.Contains("font", result.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("iwork-editable,pdf", result.ConversionEvidence!.Facts["conversionStages"]);
        Assert.Equal("0", result.ConversionEvidence.Facts["omittedSourceUnitCount"]);
        Assert.Equal("keep existing", File.ReadAllText(files.Output));
        Assert.Empty(Directory.GetFiles(files.Root, "*.tmp"));
    }
    private static void CreateNoteFreeKeynoteInput(string path) {
        IWorkKeynoteDocument source = IWorkKeynoteDocument.Create();
        source.AddSlide().AddText("hello keynote", 40, 40, 600, 180, fontSizePoints: 24);
        source.AddSlide().AddText("Second slide", 40, 40, 600, 180, fontSizePoints: 24);
        source.Save(path, overwrite: true);
    }
}
