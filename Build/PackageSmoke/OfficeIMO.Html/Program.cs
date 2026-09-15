using OfficeIMO;
using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Providers;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Word.Html;
using OfficeIMO.Pdf;

var output = new OfficeHtmlDocumentOptions {
    EmitDocumentShell = true,
    IncludeDefaultStyles = true,
    Title = "Packed API contract",
    Language = "en",
    BodyClass = "packed-consumer",
    NewLine = "\n"
};

var galleryArtifacts = new[] {
    new HtmlCapabilityGalleryArtifact("source", "html", "source.html", "text/html", 1, new string('0', 64))
};
var galleryDiagnostics = new[] {
    new HtmlDiagnostic("PackageSmoke", "Snapshot", "Packed snapshot")
};
var galleryResult = new HtmlCapabilityGalleryResult(
    new HtmlCapabilityGalleryScenario("packed", "Packed", "HTML", "Packed API proof"),
    galleryArtifacts,
    galleryDiagnostics);
PdfConversionReport ReadPdfReport(PdfHtmlConversionResult result) => result.Report;
_ = (Func<PdfHtmlConversionResult, PdfConversionReport>)ReadPdfReport;

var galleryBuilder = new HtmlCapabilityGalleryResult(
    new HtmlCapabilityGalleryScenario("builder", "Builder", "HTML", "Compatibility builder proof"));
galleryBuilder.AddArtifact(galleryArtifacts[0]);
galleryBuilder.Diagnostics.Add(galleryDiagnostics[0]);
_ = new HtmlCapabilityGalleryManifest(
    galleryBuilder,
    HtmlConversionProfile.Document,
    roundTripScore: null,
    resourceManifest: null);

WordToHtmlOptions word = WordToHtmlOptions.CreateDocumentRoundTripProfile();
word.DocumentOutput = output.Clone();
word.Profile = OfficeHtmlConversionProfile.WordDocumentRoundTrip;
ExcelHtmlSaveOptions excel = ExcelHtmlSaveOptions.CreateVisualReviewProfile();
excel.DocumentOutput = output.Clone();
excel.Profile = OfficeHtmlConversionProfile.ExcelVisualReview;
PowerPointHtmlSaveOptions powerPoint = PowerPointHtmlSaveOptions.CreateVisualReviewProfile();
powerPoint.DocumentOutput = output.Clone();
powerPoint.Profile = OfficeHtmlConversionProfile.PowerPointVisualReview;
RtfToHtmlOptions rtf = RtfToHtmlOptions.CreatePrintReviewProfile();
rtf.DocumentOutput = output.Clone();
rtf.Profile = OfficeHtmlConversionProfile.RtfPrintReview;
PdfToHtmlOptions pdf = PdfToHtmlOptions.CreatePositionedReviewProfile();
pdf.DocumentOutput = output.Clone();

if (word.SharedProfile != HtmlConversionProfile.Document ||
    excel.SharedProfile != HtmlConversionProfile.PositionedReview ||
    powerPoint.SharedProfile != HtmlConversionProfile.PositionedReview ||
    rtf.SharedProfile != HtmlConversionProfile.HighFidelityPrint) {
    throw new InvalidOperationException("The packed adapter profile mappings are inconsistent.");
}

HtmlTargetCapabilityContract pdfContract = HtmlTargetCapabilityContracts.Get(HtmlConversionTarget.Pdf);
HtmlToTargetCapabilityContract htmlToPdf = pdfContract.HtmlToTarget;
TargetToHtmlCapabilityContract pdfToHtml = pdfContract.TargetToHtml
    ?? throw new InvalidOperationException("The packed PDF-to-HTML route contract is missing.");
if (htmlToPdf.Profiles.Contains("PositionedReview", StringComparer.Ordinal) ||
    !pdfToHtml.Profiles.Contains("PositionedReview", StringComparer.Ordinal) ||
    string.IsNullOrWhiteSpace(htmlToPdf.DiagnosticsContract) ||
    string.IsNullOrWhiteSpace(pdfToHtml.DiagnosticsContract)) {
    throw new InvalidOperationException("The packed directional route contract is inconsistent.");
}

string fragment = OfficeHtmlDocumentShell.WrapBody("<p>fragment</p>", new OfficeHtmlDocumentOptions {
    EmitDocumentShell = false,
    Language = "en",
    NewLine = "\n"
});
if (!string.Equals(fragment, "<p>fragment</p>", StringComparison.Ordinal)) {
    throw new InvalidOperationException("The packed document-output fragment contract failed.");
}
if (galleryResult.Artifacts.Count != 1 || galleryResult.Diagnostics.Count != 1) {
    throw new InvalidOperationException("The packed immutable gallery-result contract failed.");
}
if (!galleryResult.IsReadOnly || !galleryResult.Diagnostics.IsReadOnly) {
    throw new InvalidOperationException("The packed gallery-result snapshot is not frozen.");
}

var parsed = HtmlDocumentEngine.Default.ParseDocument("<!DOCTYPE odd@name><h1 id='title'>Original</h1>");
var changed = parsed.Edit(edit => {
    var title = edit.QuerySelector("#title")!;
    title.TextContent = "Packed edit";
    title.SetAttribute("ID", "updated");
});
var conversion = HtmlConversionDocument.FromDocument(changed);
if (parsed.QuerySelector("#title")!.TextContent != "Original" || changed.QuerySelector("#updated") == null || !conversion.ToMarkdown().Contains("Packed edit"))
    throw new InvalidOperationException("Packed owned document/edit/Markdown contract failed.");
if (typeof(HtmlDocument).Assembly.GetReferencedAssemblies().Any(name => name.Name!.StartsWith("AngleSharp", StringComparison.Ordinal) || name.Name == "OfficeIMO.Core"))
    throw new InvalidOperationException("The owned HTML leaf references a parser or drawing implementation.");
byte[] foundationPng = conversion.ToPng();
if (foundationPng.Length < 8 || foundationPng[0] != 137 || foundationPng[1] != 80)
    throw new InvalidOperationException("Packed owned document image rendering failed.");
var foundationPdf = PdfReadDocument.Open(conversion.ToPdfBytes());
if (!foundationPdf.ExtractText().Contains("Packed edit"))
    throw new InvalidOperationException("Packed owned document PDF text was lost.");

HtmlRenderRequest imageRequest = HtmlRenderRequest.Create(
    HtmlRenderIntentProfile.ScreenFullPage,
    HtmlRenderEncoder.Png,
    new HtmlRenderOptions { ViewportWidth = 480D });
HtmlRenderResult retained = HtmlRenderEngine.Execute(conversion, imageRequest);
retained = retained.WithAdditionalDiagnostics(new[] {
    new HtmlDiagnostic("PackageSmoke", "RetainedBoundary", "Packed retained-result evidence",
        HtmlDiagnosticSeverity.Info, "package-smoke.html")
});
HtmlRenderSurface retainedSurface = retained.GetSurface(0);
if (retained.Surfaces.Count != 1 || retained.OutputSurfaces.Count != 1 ||
    retained.ExportImage().Bytes.Length < 8 || retainedSurface.CreateDrawing().Width <= 0D ||
    !retainedSurface.TryMapToSource(1D, 1D, out HtmlRenderSourcePoint? mappedSource) ||
    mappedSource == null || mappedSource.SourcePageNumber != 1 ||
    retained.Request.ProfileId != "screen-full-page-v1" ||
    !retained.DeclaredProviderIds.Contains(HtmlCapabilityProviderIds.OfficeIMOHtml))
    throw new InvalidOperationException("Packed explicit HTML render request contract failed.");
HtmlRenderArchiveResult renderArchive = retained.ExportArchive(new HtmlRenderArchiveOptions {
    MaximumArchiveBytes = 16 * 1024 * 1024
});
if (renderArchive.EncodedLength < 1 || renderArchive.Manifest.Pages.Count != 1 ||
    renderArchive.Manifest.PageSet != HtmlRenderPageSetMode.Selected ||
    renderArchive.Manifest.FirstPageIndex != 0 || renderArchive.Manifest.PageCount != 1 ||
    !renderArchive.Manifest.Diagnostics.Any(diagnostic => diagnostic.Code == "RetainedBoundary") ||
    renderArchive.Manifest.Pages[0].HasLoss != renderArchive.Manifest.Pages[0].EncodingDiagnostics.Any(
        diagnostic => diagnostic.LossKind != OfficeConversionLossKind.None) ||
    string.IsNullOrWhiteSpace(renderArchive.Manifest.Pages[0].Sha256))
    throw new InvalidOperationException("Packed HTML render archive contract failed.");

HtmlPdfRenderRequestResult explicitPdf = conversion.RenderToPdfResult(HtmlRenderRequest.Create(
    HtmlRenderIntentProfile.ScreenMediaPaged,
    HtmlRenderEncoder.Pdf,
    new HtmlToPdfOptions()));
byte[] explicitPdfBytes = explicitPdf.ToBytes();
if (explicitPdf.RenderResult.Request.CssMedia != HtmlCssMediaContext.Screen ||
    explicitPdfBytes.Length < 4 || System.Text.Encoding.ASCII.GetString(explicitPdfBytes, 0, 4) != "%PDF")
    throw new InvalidOperationException("Packed explicit HTML-to-PDF request contract failed.");

var tracedDocument = HtmlConversionDocument.Parse(
    "<style>[data-tone='IMPORTANT' i] > .notice { color:hsl(210 50 40 / 75%); opacity:calc(.2 + .3); }</style>" +
    "<section data-tone='important'><p class='notice'>Status</p></section>");
var tracedElement = tracedDocument.Document.QuerySelector(".notice")
    ?? throw new InvalidOperationException("The packed cascade-trace element was not parsed.");
HtmlComputedStyle tracedStyle = HtmlComputedStyleEngine.Compute(tracedDocument, new HtmlComputedStyleOptions {
    IncludeCascadeTraces = true
})[tracedElement];
OfficeIMO.Html.Css.HtmlCssCascadeTrace? colorTrace = tracedStyle.GetCascadeTrace("color");
if (tracedStyle.GetValue("color") != "rgba(51, 102, 153, 0.75)" || tracedStyle.GetValue("opacity") != "0.5" ||
    colorTrace?.Candidates.Count != 1 || colorTrace.Candidates[0].Decision != OfficeIMO.Html.Css.HtmlCssCascadeDecision.Selected ||
    colorTrace.Candidates[0].Source != OfficeIMO.Html.Css.HtmlCssCascadeSourceKind.StyleRule)
    throw new InvalidOperationException("The packed owned selector, typed computed-value or cascade-trace contract failed.");

Console.WriteLine("OfficeIMO HTML packed API smoke passed on " +
    System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription + ".");
