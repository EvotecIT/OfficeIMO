using System;
using System.Linq;
using System.Globalization;
using OfficeIMO.Drawing;
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
        /// Changes the stored run text casing while preserving run formatting.
        /// </summary>
        public PowerPointTextRun TransformTextCase(OfficeTextCase textCase, CultureInfo? culture = null) {
            Text = OfficeTextCaseTransformer.Apply(Text, textCase, culture);
            return this;
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
            get => UnderlineStyle is { } style && style != PowerPointUnderlineStyle.None;
            set {
                UnderlineStyle = value ? PowerPointUnderlineStyle.Single : null;
            }
        }

        /// <summary>
        /// Gets or sets the native DrawingML underline variant.
        /// </summary>
        public PowerPointUnderlineStyle? UnderlineStyle {
            get => Run.RunProperties?.Underline?.Value.ToOfficeEnum();
            set {
                A.RunProperties props = EnsureRunProperties();
                props.Underline = value?.ToOpenXml();
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is strikethrough.
        /// </summary>
        public bool Strikethrough {
            get => StrikeStyle is { } style && style != PowerPointStrikeStyle.None;
            set {
                StrikeStyle = value ? PowerPointStrikeStyle.Single : null;
            }
        }

        /// <summary>
        /// Gets or sets the native DrawingML strike-through variant.
        /// </summary>
        public PowerPointStrikeStyle? StrikeStyle {
            get {
                A.TextStrikeValues? value = Run.RunProperties?.Strike?.Value;
                if (!value.HasValue) return null;
                if (value.Value == A.TextStrikeValues.NoStrike) return PowerPointStrikeStyle.None;
                if (value.Value == A.TextStrikeValues.SingleStrike) return PowerPointStrikeStyle.Single;
                if (value.Value == A.TextStrikeValues.DoubleStrike) return PowerPointStrikeStyle.Double;
                throw new InvalidOperationException($"Unsupported DrawingML strike value '{value.Value}'.");
            }
            set {
                A.RunProperties props = EnsureRunProperties();
                props.Strike = value switch {
                    null => null,
                    PowerPointStrikeStyle.None => A.TextStrikeValues.NoStrike,
                    PowerPointStrikeStyle.Single => A.TextStrikeValues.SingleStrike,
                    PowerPointStrikeStyle.Double => A.TextStrikeValues.DoubleStrike,
                    _ => throw new ArgumentOutOfRangeException(nameof(value))
                };
            }
        }

        /// <summary>
        /// Gets or sets native DrawingML capitalization without changing the stored characters.
        /// </summary>
        public PowerPointCapitalization? Capitalization {
            get {
                A.TextCapsValues? value = Run.RunProperties?.Capital?.Value;
                if (!value.HasValue) return null;
                if (value.Value == A.TextCapsValues.None) return PowerPointCapitalization.None;
                if (value.Value == A.TextCapsValues.Small) return PowerPointCapitalization.SmallCaps;
                if (value.Value == A.TextCapsValues.All) return PowerPointCapitalization.AllCaps;
                throw new InvalidOperationException($"Unsupported DrawingML capitalization value '{value.Value}'.");
            }
            set {
                A.RunProperties props = EnsureRunProperties();
                props.Capital = value switch {
                    null => null,
                    PowerPointCapitalization.None => A.TextCapsValues.None,
                    PowerPointCapitalization.SmallCaps => A.TextCapsValues.Small,
                    PowerPointCapitalization.AllCaps => A.TextCapsValues.All,
                    _ => throw new ArgumentOutOfRangeException(nameof(value))
                };
            }
        }

        /// <summary>
        /// Gets or sets the DrawingML baseline shift in percent, from -100 through 100.
        /// Positive values create superscript and negative values create subscript.
        /// </summary>
        public double? BaselinePercent {
            get => Run.RunProperties?.Baseline?.Value is int value ? value / 1000D : null;
            set {
                if (value.HasValue && (double.IsNaN(value.Value) || double.IsInfinity(value.Value)
                    || value.Value < -100D || value.Value > 100D)) {
                    throw new ArgumentOutOfRangeException(nameof(value), "Baseline percent must be between -100 and 100.");
                }

                A.RunProperties props = EnsureRunProperties();
                props.Baseline = value.HasValue
                    ? checked((int)Math.Round(value.Value * 1000D, MidpointRounding.AwayFromZero))
                    : null;
            }
        }

        /// <summary>Applies superscript using a 30 percent baseline shift.</summary>
        public PowerPointTextRun SetSuperscript(double baselinePercent = 30D) {
            BaselinePercent = Math.Abs(baselinePercent);
            return this;
        }

        /// <summary>Applies subscript using a 25 percent baseline shift.</summary>
        public PowerPointTextRun SetSubscript(double baselinePercent = 25D) {
            BaselinePercent = -Math.Abs(baselinePercent);
            return this;
        }

        /// <summary>Restores the run to the normal text baseline.</summary>
        public PowerPointTextRun SetBaseline() {
            BaselinePercent = null;
            return this;
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
                FontSizePoints = value;
            }
        }

        /// <summary>
        /// Gets or sets the font size in points while preserving DrawingML hundredth-point precision.
        /// </summary>
        public double? FontSizePoints {
            get {
                int? size = Run.RunProperties?.FontSize?.Value;
                return size.HasValue ? size.Value / 100D : (double?)null;
            }
            set {
                A.RunProperties props = EnsureRunProperties();
                props.FontSize = PowerPointTextDefaults.ToDrawingFontSize(value, nameof(value));
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
            OpenXmlPart? ownerPart = _ownerPart as OpenXmlPart ?? _slidePart;
            if (ownerPart == null) {
                throw new InvalidOperationException("Hyperlinks require an owning presentation part.");
            }

            HyperlinkRelationship rel = ownerPart.AddHyperlinkRelationship(uri, true);
            A.RunProperties props = EnsureRunProperties();
            var hyperlink = new A.HyperlinkOnClick { Id = rel.Id };
            if (!string.IsNullOrWhiteSpace(tooltip)) {
                hyperlink.Tooltip = tooltip;
            }
            ReplaceClickHyperlink(props, hyperlink);
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
            OpenXmlPart ownerPart = _ownerPart as OpenXmlPart ?? _slidePart;

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

            string relationshipId;
            if (ownerPart is NotesSlidePart
                && !ReferenceEquals(_slidePart, targetSlide.SlidePart)) {
                Uri targetUri = PowerPointHyperlinkResolver.CreatePartRelativeUri(
                    ownerPart, targetSlide.SlidePart);
                HyperlinkRelationship relationship = ownerPart.HyperlinkRelationships
                    .FirstOrDefault(candidate => !candidate.IsExternal
                        && candidate.Uri == targetUri)
                    ?? ownerPart.AddHyperlinkRelationship(targetUri, false);
                relationshipId = relationship.Id;
            } else {
                if (!ownerPart.Parts.Any(pair => ReferenceEquals(
                        pair.OpenXmlPart, targetSlide.SlidePart))) {
                    ownerPart.AddPart(targetSlide.SlidePart);
                }
                relationshipId = ownerPart.GetIdOfPart(targetSlide.SlidePart);
            }

            A.RunProperties props = EnsureRunProperties();
            var hyperlink = new A.HyperlinkOnClick {
                Id = relationshipId,
                Action = "ppaction://hlinksldjump"
            };
            if (!string.IsNullOrWhiteSpace(tooltip)) {
                hyperlink.Tooltip = tooltip;
            }
            ReplaceClickHyperlink(props, hyperlink);
        }

        /// <summary>
        /// Removes any hyperlink from this run.
        /// </summary>
        public void ClearHyperlink() {
            A.RunProperties? props = Run.RunProperties;
            if (props != null) ReplaceClickHyperlink(props, replacement: null);
        }

        private void ReplaceClickHyperlink(A.RunProperties properties,
            A.HyperlinkOnClick? replacement) {
            A.HyperlinkOnClick[] previous = properties
                .Elements<A.HyperlinkOnClick>().ToArray();
            if (replacement != null) {
                A.HyperlinkSound? preservedSound = previous
                    .SelectMany(link => link.Elements<A.HyperlinkSound>())
                    .FirstOrDefault();
                if (preservedSound != null) {
                    replacement.Append((A.HyperlinkSound)preservedSound
                        .CloneNode(true));
                }
                bool? preservedEndSound = previous
                    .Select(link => link.EndSound?.Value)
                    .FirstOrDefault(value => value.HasValue);
                if (preservedEndSound.HasValue) {
                    replacement.EndSound = preservedEndSound.Value;
                }
            }
            string[] relationshipIds = previous
                .Select(link => link.Id?.Value)
                .Where(id => !string.IsNullOrEmpty(id))
                .Cast<string>()
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            string[] soundRelationshipIds = previous
                .SelectMany(link => link.Elements<A.HyperlinkSound>())
                .Select(sound => sound.Embed?.Value)
                .Where(id => !string.IsNullOrEmpty(id))
                .Cast<string>()
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            foreach (A.HyperlinkOnClick hyperlink in previous) {
                hyperlink.Remove();
            }
            if (replacement != null) properties.Append(replacement);
            OpenXmlPart? ownerPart = _ownerPart as OpenXmlPart ?? _slidePart;
            if (ownerPart == null) return;
            foreach (string relationshipId in relationshipIds) {
                RemoveHyperlinkRelationshipIfUnused(ownerPart,
                    relationshipId, _slidePart);
            }
            foreach (string soundRelationshipId in soundRelationshipIds) {
                PowerPointEmbeddedSound.RemoveIfUnused(ownerPart,
                    soundRelationshipId);
            }
        }

        private static void RemoveHyperlinkRelationshipIfUnused(
            OpenXmlPart ownerPart, string relationshipId,
            SlidePart? owningSlidePart) {
            if (ReferencesRelationship(ownerPart.RootElement,
                    relationshipId)) return;
            if (ownerPart is NotesSlidePart
                && owningSlidePart != null
                && ownerPart.Parts.Any(pair => string.Equals(
                        pair.RelationshipId, relationshipId,
                        StringComparison.Ordinal)
                    && ReferenceEquals(pair.OpenXmlPart,
                        owningSlidePart))) return;
            HyperlinkRelationship? external = ownerPart
                .HyperlinkRelationships.FirstOrDefault(relationship =>
                    string.Equals(relationship.Id, relationshipId,
                        StringComparison.Ordinal));
            if (external != null) {
                ownerPart.DeleteReferenceRelationship(external);
                return;
            }
            if (ownerPart.Parts.Any(pair => string.Equals(
                    pair.RelationshipId, relationshipId,
                    StringComparison.Ordinal))) {
                ownerPart.DeletePart(relationshipId);
            }
        }

        private static bool ReferencesRelationship(
            OpenXmlPartRootElement? root,
            string relationshipId) => root != null
            && (root.GetAttributes().Any(attribute => string.Equals(
                    attribute.NamespaceUri,
                    PowerPointUtils.RelationshipIdNamespace,
                    StringComparison.Ordinal)
                && string.Equals(attribute.Value, relationshipId,
                    StringComparison.Ordinal))
                || root.Descendants().Any(element => element
                    .GetAttributes().Any(attribute => string.Equals(
                            attribute.NamespaceUri,
                            PowerPointUtils.RelationshipIdNamespace,
                            StringComparison.Ordinal)
                        && string.Equals(attribute.Value, relationshipId,
                            StringComparison.Ordinal))));

        private A.RunProperties EnsureRunProperties() {
            return Run.RunProperties ??= new A.RunProperties();
        }
    }
}
