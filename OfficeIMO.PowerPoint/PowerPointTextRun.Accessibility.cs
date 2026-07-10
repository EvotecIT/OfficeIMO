using System;
using DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTextRun {
        /// <summary>Gets or sets the BCP 47 language tag applied to this text run.</summary>
        public string? Language {
            get => Run.RunProperties?.Language?.Value;
            set {
                A.RunProperties properties = EnsureRunProperties();
                properties.Language = NormalizeLanguage(value);
            }
        }

        /// <summary>Gets or sets the accessible tooltip associated with this run's click hyperlink.</summary>
        public string? HyperlinkTooltip {
            get => Run.RunProperties?.GetFirstChild<A.HyperlinkOnClick>()?.Tooltip?.Value;
            set {
                A.HyperlinkOnClick? hyperlink = Run.RunProperties?.GetFirstChild<A.HyperlinkOnClick>();
                if (hyperlink == null) {
                    if (value == null) return;
                    throw new InvalidOperationException("A hyperlink must be assigned before setting its tooltip.");
                }
                hyperlink.Tooltip = string.IsNullOrWhiteSpace(value) ? null : value!.Trim();
            }
        }

        /// <summary>Returns whether the visible link text is meaningful without surrounding context.</summary>
        public bool HasMeaningfulHyperlinkLabel => Hyperlink == null || IsMeaningfulLinkLabel(Text);

        internal static bool IsMeaningfulLinkLabel(string? value) {
            if (string.IsNullOrWhiteSpace(value)) return false;
            string normalized = value!.Trim();
            if (normalized.Length < 3) return false;
            if (Uri.TryCreate(normalized, UriKind.Absolute, out _)) return false;
            return !string.Equals(normalized, "click here", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(normalized, "here", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(normalized, "more", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(normalized, "link", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(normalized, "read more", StringComparison.OrdinalIgnoreCase);
        }

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
