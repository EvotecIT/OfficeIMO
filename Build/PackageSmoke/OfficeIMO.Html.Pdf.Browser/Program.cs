using HtmlTinkerX;
using OfficeIMO.Html.Pdf.Browser;

var policy = HtmlBrowserNetworkPolicy.Offline;
var options = new HtmlBrowserPdfRendererOptions(
    viewportWidth: 390,
    viewportHeight: 844,
    deviceScaleFactor: 3,
    isMobile: true,
    hasTouch: true,
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
