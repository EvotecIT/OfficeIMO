using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTextRun {
        /// <summary>Gets or sets the BCP 47 language tag applied to this text run.</summary>
        public string? Language {
            get => PropertiesElement.Attribute("lang");
            set => SetProperty("lang", NormalizeLanguage(value));
        }

        /// <summary>Gets or sets the accessible tooltip associated with this run's click hyperlink.</summary>
        public string? HyperlinkTooltip {
            get => ClickHyperlinkElement.Attribute("tooltip");
            set {
                PowerPointXmlNode hyperlink = ClickHyperlinkElement;
                if (hyperlink.IsNone) {
                    if (value == null) return;
                    throw new InvalidOperationException("A hyperlink must be assigned before setting its tooltip.");
                }
                hyperlink.SetAttribute("tooltip", string.IsNullOrWhiteSpace(value) ? null : value!.Trim());
                MarkEdited();
                SyncLineContinuation();
            }
        }

        /// <summary>Returns whether the visible link text is meaningful without surrounding context.</summary>
        public bool HasMeaningfulHyperlinkLabel => Hyperlink == null || IsMeaningfulLinkLabel(Text);

        internal static bool IsMeaningfulLinkLabel(string? value) => OfficeIMO.Core.Internal.OfficeAccessibilityText.IsMeaningfulLinkLabel(value);

        internal static string? NormalizeLanguage(string? value) {
            if (string.IsNullOrWhiteSpace(value)) return null;
            string normalized = value!.Trim();
            if (normalized.Length > 85 || normalized.IndexOfAny(new[] { ' ', '_', '\t', '\r', '\n' }) >= 0) {
                throw new ArgumentException("Language must be a BCP 47 tag such as 'en-US'.", nameof(value));
            }
            return normalized;
        }
    }
}
