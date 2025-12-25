using System;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents a formatted text run within a paragraph.
    /// </summary>
    public class PowerPointTextRun {
        private readonly SlidePart? _slidePart;

        internal PowerPointTextRun(A.Run run, SlidePart? slidePart = null) {
            Run = run;
            _slidePart = slidePart;
        }

        internal A.Run Run { get; }

        /// <summary>
        /// Text content of the run.
        /// </summary>
        public string Text {
            get => Run.Text?.Text ?? string.Empty;
            set {
                Run.Text ??= new A.Text();
                Run.Text.Text = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is bold.
        /// </summary>
        public bool Bold {
            get => Run.RunProperties?.Bold?.Value == true;
            set {
                A.RunProperties props = EnsureRunProperties();
                props.Bold = value ? true : null;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is italic.
        /// </summary>
        public bool Italic {
            get => Run.RunProperties?.Italic?.Value == true;
            set {
                A.RunProperties props = EnsureRunProperties();
                props.Italic = value ? true : null;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is underlined.
        /// </summary>
        public bool Underline {
            get => Run.RunProperties?.Underline?.Value == A.TextUnderlineValues.Single;
            set {
                A.RunProperties props = EnsureRunProperties();
                props.Underline = value ? A.TextUnderlineValues.Single : null;
            }
        }

        /// <summary>
        /// Gets or sets the font size in points.
        /// </summary>
        public int? FontSize {
            get {
                int? size = Run.RunProperties?.FontSize?.Value;
                return size != null ? size / 100 : null;
            }
            set {
                A.RunProperties props = EnsureRunProperties();
                props.FontSize = value != null ? value * 100 : null;
            }
        }

        /// <summary>
        /// Gets or sets the font name (Latin).
        /// </summary>
        public string? FontName {
            get => Run.RunProperties?.GetFirstChild<A.LatinFont>()?.Typeface;
            set {
                A.RunProperties props = EnsureRunProperties();
                props.RemoveAllChildren<A.LatinFont>();
                if (value != null) {
                    props.Append(new A.LatinFont { Typeface = value });
                }
            }
        }

        /// <summary>
        /// Gets or sets the text color in hexadecimal format (e.g. "FF0000").  
        /// </summary>
        public string? Color {
            get => Run.RunProperties?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val;
            set {
                A.RunProperties props = EnsureRunProperties();
                var latin = props.GetFirstChild<A.LatinFont>();
                var ea = props.GetFirstChild<A.EastAsianFont>();
                var cs = props.GetFirstChild<A.ComplexScriptFont>();

                props.RemoveAllChildren<A.SolidFill>();
                props.RemoveAllChildren<A.LatinFont>();
                props.RemoveAllChildren<A.EastAsianFont>();
                props.RemoveAllChildren<A.ComplexScriptFont>();

                if (value != null) {
                    props.Append(new A.SolidFill(new A.RgbColorModelHex { Val = value }));
                }

                if (latin != null) props.Append((A.LatinFont)latin.CloneNode(true));
                if (ea != null) props.Append((A.EastAsianFont)ea.CloneNode(true));
                if (cs != null) props.Append((A.ComplexScriptFont)cs.CloneNode(true));
            }
        }

        /// <summary>
        /// Gets or sets the highlight color in hexadecimal format (e.g. "FFFF00").
        /// </summary>
        public string? HighlightColor {
            get => Run.RunProperties?.GetFirstChild<A.Highlight>()?.GetFirstChild<A.RgbColorModelHex>()?.Val;
            set {
                A.RunProperties props = EnsureRunProperties();
                props.RemoveAllChildren<A.Highlight>();
                if (value != null) {
                    props.Append(new A.Highlight(new A.RgbColorModelHex { Val = value }));
                }
            }
        }

        /// <summary>
        /// Gets or sets the hyperlink target for this run.
        /// </summary>
        public Uri? Hyperlink {
            get {
                if (_slidePart == null) {
                    return null;
                }

                string? relId = Run.RunProperties?.GetFirstChild<A.HyperlinkOnClick>()?.Id;
                if (string.IsNullOrWhiteSpace(relId)) {
                    return null;
                }

                HyperlinkRelationship? rel = _slidePart.HyperlinkRelationships
                    .FirstOrDefault(r => string.Equals(r.Id, relId, StringComparison.Ordinal));
                return rel?.Uri;
            }
            set {
                if (value == null) {
                    ClearHyperlink();
                } else {
                    SetHyperlink(value);
                }
            }
        }

        /// <summary>
        /// Sets a hyperlink for this run.
        /// </summary>
        public void SetHyperlink(string url, string? tooltip = null) {
            if (url == null) {
                throw new ArgumentNullException(nameof(url));
            }

            SetHyperlink(new Uri(url, UriKind.RelativeOrAbsolute), tooltip);
        }

        /// <summary>
        /// Sets a hyperlink for this run.
        /// </summary>
        public void SetHyperlink(Uri uri, string? tooltip = null) {
            if (uri == null) {
                throw new ArgumentNullException(nameof(uri));
            }
            if (_slidePart == null) {
                throw new InvalidOperationException("Hyperlinks require a slide context.");
            }

            HyperlinkRelationship rel = _slidePart.AddHyperlinkRelationship(uri, true);
            A.RunProperties props = EnsureRunProperties();
            props.RemoveAllChildren<A.HyperlinkOnClick>();
            var hyperlink = new A.HyperlinkOnClick { Id = rel.Id };
            if (!string.IsNullOrWhiteSpace(tooltip)) {
                hyperlink.Tooltip = tooltip;
            }
            props.Append(hyperlink);
        }

        /// <summary>
        /// Removes any hyperlink from this run.
        /// </summary>
        public void ClearHyperlink() {
            A.RunProperties? props = Run.RunProperties;
            props?.RemoveAllChildren<A.HyperlinkOnClick>();
        }

        private A.RunProperties EnsureRunProperties() {
            return Run.RunProperties ??= new A.RunProperties();
        }
    }
}
