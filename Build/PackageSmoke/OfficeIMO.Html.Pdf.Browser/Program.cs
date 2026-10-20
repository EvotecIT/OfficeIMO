using HtmlTinkerX;
using OfficeIMO.Html.Pdf.Browser;
using OfficeIMO.Pdf;

if (args.Length == 1 && args[0] == "--install-browser") {
    await HtmlBrowser.EnsureInstalledAsync(HtmlBrowserEngine.Chromium);
    Console.WriteLine("HtmlTinkerX Chromium provisioning passed.");
    return;
}
if (args.Length != 0 && (args.Length != 2 || args[0] != "--capture"))
    throw new ArgumentException("Use --install-browser, --capture <output.pdf>, or no arguments for the API smoke.");

var policy = HtmlBrowserNetworkPolicy.Offline;
var options = new HtmlBrowserPdfRendererOptions(
    new HtmlBrowserPdfDeviceEmulation(
        deviceScaleFactor: 3,
        isMobile: true,
        hasTouch: true),
    viewportWidth: 390,
    viewportHeight: 844,
    networkPolicy: policy);

if (policy.AllowNetworkAccess ||
    options.DeviceScaleFactor != 3 ||
    options.IsMobile != true ||
    options.HasTouch != true ||
    !ReferenceEquals(options.NetworkPolicy, policy)) {
    throw new InvalidOperationException(
        "The packed HtmlTinkerX browser PDF device or offline policy contract is inconsistent.");
}

bool hasOfficeBridge = typeof(HtmlBrowserPdfOfficeExtensions)
    .GetMethods()
    .Any(method => string.Equals(
        method.Name,
        nameof(HtmlBrowserPdfOfficeExtensions.CapturePdfDocumentResultAsync),
        StringComparison.Ordinal));
if (!hasOfficeBridge) {
    throw new InvalidOperationException(
        "The packed OfficeIMO browser PDF bridge is missing its diagnostic-preserving capture API.");
}

Console.WriteLine(
    "OfficeIMO browser PDF packed API smoke passed on " +
    System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription + ".");

if (args.Length == 2) {
    await using var renderer = new HtmlBrowserPdfRenderer(options);
    var request = new HtmlBrowserPdfRequest(
        HtmlBrowserPdfSource.FromHtml("""
            <!doctype html><html lang="en"><head><title>Browser accessibility smoke</title></head>
            <body><h1 id="heading">Waiting for JavaScript</h1><p>Offline browser capture.</p>
            <script>document.getElementById('heading').textContent = 'JavaScript capture passed';</script>
            </body></html>
            """),
        pdfOptions: new HtmlBrowserPdfOptions(tagged: true));
    PdfDocumentConversionResult result = await renderer.CapturePdfDocumentResultAsync(request);
    byte[] pdf = result.ToBytes();
    PdfReadDocument inspected = PdfReadDocument.Open(pdf);
    if (!inspected.HasTaggedContent ||
        inspected.ExtractText().IndexOf("JavaScript capture passed", StringComparison.Ordinal) < 0)
        throw new InvalidOperationException("Browser capture did not retain its executed JavaScript and tagged PDF structure.");
    File.WriteAllBytes(args[1], pdf);
    Console.WriteLine("Tagged browser PDF capture passed: " + Path.GetFullPath(args[1]));
}
