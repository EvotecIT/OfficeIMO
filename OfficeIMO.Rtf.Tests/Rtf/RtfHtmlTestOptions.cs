using OfficeIMO.Html;

namespace OfficeIMO.Tests.Rtf;

internal static class RtfHtmlTestOptions {
    internal static RtfHtmlExportOptions CreateRoundTripFragment() {
        RtfHtmlExportOptions options = RtfHtmlExportOptions.CreateRoundTripProfile();
        options.FragmentOnly = true;
        return options;
    }
}
