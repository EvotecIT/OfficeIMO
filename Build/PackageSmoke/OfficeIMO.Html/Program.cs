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
PdfHtmlSaveOptions pdf = PdfHtmlSaveOptions.CreatePositionedReviewProfile();
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

Console.WriteLine("OfficeIMO HTML packed API smoke passed on " +
    System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription + ".");
