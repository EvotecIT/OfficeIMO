using OfficeIMO.Html;

namespace OfficeIMO.Tests.Rtf;

internal static class RtfHtmlTestOptions {
    internal static RtfToHtmlOptions CreateRoundTripFragment() {
        RtfToHtmlOptions options = RtfToHtmlOptions.CreateRoundTripProfile();
        options.FragmentOnly = true;
        return options;
    }
}
