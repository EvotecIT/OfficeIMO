using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointCustomShowAction {
        internal const string Prefix = "ppaction://customshow?id=";
        private const string ReturnSuffix = "&return=true";

        internal static bool IsCustomShowAction(string? action) =>
            action?.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase) == true;

        internal static bool TryParseSupported(string? action, out uint customShowId,
            out bool returnsToSlide) {
            customShowId = 0;
            returnsToSlide = false;
            if (!IsCustomShowAction(action)) return false;
            string value = action!.Substring(Prefix.Length);
            if (value.EndsWith(ReturnSuffix, StringComparison.OrdinalIgnoreCase)) {
                returnsToSlide = true;
                value = value.Substring(0, value.Length - ReturnSuffix.Length);
            }
            return value.Length > 0 && value.IndexOf('&') < 0
                && uint.TryParse(value, NumberStyles.None,
                    CultureInfo.InvariantCulture, out customShowId);
        }

        internal static bool TryValidateSupportedHyperlink(
            A.HyperlinkType hyperlink,
            out uint customShowId,
            out bool returnsToSlide,
            out string? reason) {
            if (hyperlink == null) throw new ArgumentNullException(nameof(hyperlink));
            customShowId = 0;
            returnsToSlide = false;
            reason = null;
            if (!TryParseSupported(hyperlink.Action?.Value,
                    out customShowId, out returnsToSlide)) {
                reason = "The custom-show action is malformed.";
                return false;
            }
            if (!string.IsNullOrEmpty(hyperlink.Id?.Value)
                || !string.IsNullOrEmpty(hyperlink.Tooltip?.Value)) {
                reason = "Custom-show actions cannot combine a relationship or screen tip.";
                return false;
            }

            A.HyperlinkSound[] sounds = hyperlink.Elements<A.HyperlinkSound>()
                .ToArray();
            if (sounds.Length > 1
                || hyperlink.ChildElements.Any(child =>
                    child is not A.HyperlinkSound)
                || hyperlink.GetAttributes().Any(attribute =>
                    !IsSupportedHyperlinkAttribute(attribute))) {
                reason = "Custom-show actions contain unsupported hyperlink attributes or child markup.";
                return false;
            }
            return true;
        }

        private static bool IsSupportedHyperlinkAttribute(
            OpenXmlAttribute attribute) =>
            string.Equals(attribute.LocalName, "id", StringComparison.Ordinal)
            || string.Equals(attribute.LocalName, "action", StringComparison.Ordinal)
            || string.Equals(attribute.LocalName, "tooltip", StringComparison.Ordinal)
            || string.Equals(attribute.LocalName, "highlightClick", StringComparison.Ordinal)
            || string.Equals(attribute.LocalName, "endSnd", StringComparison.Ordinal);
    }
}
