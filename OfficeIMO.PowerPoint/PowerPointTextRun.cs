using System;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents a formatted text run within a paragraph.
    /// </summary>
    public partial class PowerPointTextRun {
        private readonly SlidePart? _slidePart;
        private readonly OpenXmlPartContainer? _ownerPart;

        internal PowerPointTextRun(A.Run run, SlidePart? slidePart = null, OpenXmlPartContainer? ownerPart = null) {
            Run = run;
            _slidePart = slidePart;
            _ownerPart = ownerPart ?? slidePart;
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
        /// Gets or sets a value indicating whether the run is strikethrough.
        /// </summary>
        public bool Strikethrough {
            get => Run.RunProperties?.Strike?.Value == A.TextStrikeValues.SingleStrike;
            set {
                A.RunProperties props = EnsureRunProperties();
                props.Strike = value ? A.TextStrikeValues.SingleStrike : null;
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
        /// Gets or sets the hyperlink target for this run. Internal slide links are returned as
        /// stable Markdown-compatible fragments such as <c>#slide-2</c>.
        /// </summary>
        public Uri? Hyperlink {
            get {
                if (_ownerPart == null) {
                    return null;
                }

                return PowerPointHyperlinkResolver.Resolve(_ownerPart,
                    _slidePart, Run.RunProperties?
                        .GetFirstChild<A.HyperlinkOnClick>());
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
        /// Sets an internal hyperlink from this run to another slide in the same presentation.
        /// </summary>
        public void SetHyperlink(PowerPointSlide targetSlide,
            string? tooltip = null) {
            if (targetSlide == null) {
                throw new ArgumentNullException(nameof(targetSlide));
            }
            if (_slidePart == null) {
                throw new InvalidOperationException(
                    "Hyperlinks require a slide context.");
            }

            PresentationPart? sourcePresentation = _slidePart.GetParentParts()
                .OfType<PresentationPart>().FirstOrDefault();
            PresentationPart? targetPresentation = targetSlide.SlidePart
                .GetParentParts().OfType<PresentationPart>().FirstOrDefault();
            if (sourcePresentation == null
                || !ReferenceEquals(sourcePresentation, targetPresentation)) {
                throw new ArgumentException(
                    "The hyperlink target must belong to the same presentation.",
                    nameof(targetSlide));
            }

            if (!_slidePart.Parts.Any(pair => ReferenceEquals(
                    pair.OpenXmlPart, targetSlide.SlidePart))) {
                _slidePart.AddPart(targetSlide.SlidePart);
            }

            A.RunProperties props = EnsureRunProperties();
            props.RemoveAllChildren<A.HyperlinkOnClick>();
            var hyperlink = new A.HyperlinkOnClick {
                Id = _slidePart.GetIdOfPart(targetSlide.SlidePart),
                Action = "ppaction://hlinksldjump"
            };
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
