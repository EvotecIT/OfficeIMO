using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static RtfHtmlEncapsulation? ReadHtmlEncapsulation(RtfGroup root, int ansiCodePage, int unicodeSkipCount) {
            RtfControlWord? fromHtml = root.Children
                .OfType<RtfControlWord>()
                .FirstOrDefault(control => control.Name == "fromhtml");
            if (fromHtml == null) return null;

            var html = new StringBuilder();
            AppendHtmlDestinations(root, html, ansiCodePage, unicodeSkipCount);
            return new RtfHtmlEncapsulation(fromHtml.Parameter ?? 1, html.ToString());
        }

        private static void AppendHtmlDestinations(RtfGroup group, StringBuilder html, int ansiCodePage, int unicodeSkipCount) {
            foreach (RtfNode node in group.Children) {
                if (!(node is RtfGroup childGroup)) continue;
                if (childGroup.Destination == "htmltag" || childGroup.Destination == "mhtmltag") {
                    html.Append(CollectPlainText(childGroup, ansiCodePage, unicodeSkipCount));
                } else {
                    AppendHtmlDestinations(childGroup, html, ansiCodePage, unicodeSkipCount);
                }
            }
        }
    }
}
