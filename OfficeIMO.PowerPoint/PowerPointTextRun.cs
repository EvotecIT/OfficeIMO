using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents a formatted text run within a paragraph.
    /// </summary>
    public partial class PowerPointTextRun : IPackageSdkBinding {
        // Children of a:r and of a:fld, in sequence order.
        private static readonly OfficeXmlChildOrder RunOrder = new(
            new[] { OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML }, new[] { "rPr", "t" });
        private static readonly OfficeXmlChildOrder FieldOrder = new(
            new[] { OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML }, new[] { "rPr", "pPr", "t" });

        private OfficeXmlNode _element;
        private readonly bool _isField;
        private readonly PowerPointPartRef _initialSlidePart;
        private readonly PowerPointPartRef _initialOwnerPart;
        private readonly PowerPointSlide? _packageSlide;
        // The slide whose part (its own or its notes) holds the run, when known: once removed, nothing binds.
        private readonly PowerPointSlide? _owningSlide;
        // Whether the element is a standalone copy of one removed from its part.
        private bool _standalone;
        private readonly PowerPointPartRef _packagePart;
        private List<OfficeXmlNode>? _lineContinuation;

        /// <summary>
        /// A run (<c>a:r</c>) or a dynamic field (<c>a:fld</c>) of a part, wherever the part is held.
        /// <paramref name="slidePart"/> is the slide the text belongs to and <paramref name="ownerPart"/> the part that
        /// holds it, such as a notes slide or a layout.
        /// </summary>
        internal PowerPointTextRun(OfficeXmlNode element, PowerPointPartRef slidePart, PowerPointPartRef ownerPart, List<OfficeXmlNode>? lineContinuation = null,
            PowerPointSlide? owningSlide = null) {
            _element = element;
            _owningSlide = owningSlide;
            _isField = element.Is(OfficeXmlNamespaces.DrawingML, "fld");
            _initialSlidePart = !slidePart.IsNone ? slidePart : ownerPart.Kind == PowerPointPartKind.Slide ? ownerPart : default;
            _initialOwnerPart = !ownerPart.IsNone ? ownerPart : slidePart;
            if (lineContinuation is { Count: > 0 }) _lineContinuation = lineContinuation;
            // A run of a part's package tree binds to its SDK element when the SDK opens.
            if (element.Tree != null && !_initialOwnerPart.IsNone && !_initialOwnerPart.IsSdk && element.IsAttached) {
                _packagePart = _initialOwnerPart;
                _initialOwnerPart.Document?.RegisterPackageBinding(this);
            }
        }

        // A run of a slide's package tree; it binds to the same run in the SDK when the SDK opens.
        internal PowerPointTextRun(PowerPointSlide slide, OfficeXmlTree tree, int node) {
            _packageSlide = slide;
            _packagePart = slide.PartRef;
            _element = new OfficeXmlNode(tree, node);
            _isField = _element.Is(OfficeXmlNamespaces.DrawingML, "fld");
            _initialSlidePart = _initialOwnerPart = _packagePart;
            slide.OwnerDocument?.RegisterPackageBinding(this);
        }

        /// <summary>The run's element as it is held now: once the SDK opened, a run of a package tree is its SDK element.</summary>
        /// <summary>
        /// Whether the run is still in <paramref name="paragraph"/>. A package run that was removed answers without
        /// binding, which would fail once the SDK is open.
        /// </summary>
        internal bool BelongsTo(PowerPointParagraph paragraph) {
            if (_element.Tree != null && !_element.IsAttached) return false;
            return Element.IsInside(paragraph.Element);
        }

        internal OfficeXmlNode Element {
            get {
                // A removed slide's part name can be reused by a new slide, so its content never binds.
                if (!_standalone && _element.Tree != null && _element.IsAttached && !_packagePart.IsNone && _packagePart.Document is { IsSdkLoaded: true } && _packageSlide is not { RemovedFromPackage: true }
                    && _owningSlide is not { RemovedFromPackage: true }) {
                    // The lines the run continues on bind with it, so formatting keeps reaching them.
                    if (_lineContinuation != null) {
                        for (int index = 0; index < _lineContinuation.Count; index++) {
                            OfficeXmlNode line = _lineContinuation[index];
                            // A line copied out when it was removed stays with its copy.
                            if (ReferenceEquals(line.Tree, _element.Tree)) _lineContinuation[index] = OfficeXmlNode.FromSdk(PowerPointSlide.BindPackageNode(_packagePart, line.Tree, line.Node));
                        }
                    }
                    _element = OfficeXmlNode.FromSdk(PowerPointSlide.BindPackageNode(_packagePart, _element.Tree, _element.Node));
                }
                return _element;
            }
        }

        void IPackageSdkBinding.KeepDetached() {
            // The lines the run continues on are copied with it, so its text and formatting stay whole.
            if (_lineContinuation != null) {
                for (int index = 0; index < _lineContinuation.Count; index++) {
                    OfficeXmlNode line = _lineContinuation[index];
                    if (line.Tree != null && !line.IsAttached) _lineContinuation[index] = PowerPointHeldNode.Standalone(line);
                }
            }
            if (!_standalone && _element.Tree != null && !_element.IsAttached) {
                _element = PowerPointHeldNode.Standalone(_element);
                _standalone = true;
            }
        }

        void IPackageSdkBinding.BindSdk() {
            try {
                _ = Element;
            } catch (InvalidOperationException) {
                // A run the opened slide doesn't have where expected reports it when it is used.
            }
        }

        /// <summary>The slide the run's text is on, found from where the element is now when it can be.</summary>
        internal PowerPointPartRef SlidePartRef {
            get {
                PowerPointPartRef current = ContainingPart(Element);
                return current.Kind == PowerPointPartKind.Slide ? current : _initialSlidePart;
            }
        }

        /// <summary>The part that holds the run, found from where the element is now when it can be.</summary>
        internal PowerPointPartRef OwnerPartRef {
            get {
                PowerPointPartRef current = ContainingPart(Element);
                return current.Kind == PowerPointPartKind.Slide ? current : !_initialOwnerPart.IsNone ? _initialOwnerPart : _initialSlidePart;
            }
        }

        // The part an SDK element sits in; a tree node's part is the one the run was made for.
        private static PowerPointPartRef ContainingPart(OfficeXmlNode element) =>
            element.IsSdk ? PowerPointPartRef.FromSdk(OfficeSdkXml.PartOf(element.SdkElement!)) : default;

        private PowerPointSlide? Slide => _packageSlide ?? FindSlide(SlidePartRef);

        private static PowerPointSlide? FindSlide(PowerPointPartRef part) {
            if (part.IsNone || part.Document is not { } document) return null;
            foreach (PowerPointSlide slide in document.Slides) {
                if (slide.PartRef.Equals(part)) return slide;
            }
            return null;
        }

        private OfficeXmlChildOrder Order => _isField ? FieldOrder : RunOrder;

        /// <summary>The run's <c>a:rPr</c>, or none.</summary>
        internal OfficeXmlNode PropertiesElement => Element.Child(OfficeXmlNamespaces.DrawingML, "rPr");

        /// <summary>The run's <c>a:rPr</c>, added first in the run when it has none.</summary>
        internal OfficeXmlNode EnsurePropertiesElement() {
            OfficeXmlNode element = Element;
            OfficeXmlNode properties = element.GetOrAddChild(OfficeXmlNamespaces.DrawingML, "rPr", Order);
            MarkEdited();
            return properties;
        }

        private void MarkEdited() => OwnerPartRef.MarkEdited();

        /// <summary>The last run of the text this wrapper wrote, which is the run itself for a single line.</summary>
        internal OfficeXmlNode LastLineElement {
            get {
                if (_lineContinuation != null) {
                    for (int index = _lineContinuation.Count - 1; index >= 0; index--) {
                        if (_lineContinuation[index].Is(OfficeXmlNamespaces.DrawingML, "r")) return _lineContinuation[index];
                    }
                }
                return Element;
            }
        }

        /// <summary>
        /// Text content of the run. A line break in the assigned text continues the run with
        /// <c>a:br</c> and a further run per line; this instance keeps formatting all of those lines.
        /// A dynamic field holds one line, so its breaks become spaces.
        /// </summary>
        public string Text {
            get {
                string text = Element.ChildText(OfficeXmlNamespaces.DrawingML, "t") ?? string.Empty;
                if (_lineContinuation == null) return text;
                var lines = new List<string> { text };
                foreach (OfficeXmlNode line in _lineContinuation) {
                    if (line.Is(OfficeXmlNamespaces.DrawingML, "r")) lines.Add(line.ChildText(OfficeXmlNamespaces.DrawingML, "t") ?? string.Empty);
                }
                return string.Join(Environment.NewLine, lines);
            }
            set {
                OfficeXmlNode element = Element;
                if (_isField) {
                    element.SetChildText(OfficeXmlNamespaces.DrawingML, "t", PowerPointTextContent.ToSingleLine(value), FieldOrder);
                    MarkEdited();
                    return;
                }
                if (_lineContinuation != null) {
                    foreach (OfficeXmlNode line in _lineContinuation) {
                        if (!line.Parent.IsNone) line.Remove();
                    }
                    _lineContinuation = null;
                }
                string[] lines = PowerPointTextContent.SplitLines(value);
                OfficeXmlNode text = element.GetOrAddChild(OfficeXmlNamespaces.DrawingML, "t", RunOrder);
                OfficeXmlNode parent = element.Parent;
                if (parent.IsNone) {
                    text.SetInnerText(PowerPointTextContent.ToSingleLine(value));
                    MarkEdited();
                    return;
                }
                text.SetInnerText(lines[0]);
                if (lines.Length > 1) {
                    // The further lines follow the run, each with a copy of its properties.
                    string rest = string.Join("\n", lines, 1, lines.Length - 1);
                    var continuation = new List<OfficeXmlNode>();
                    OfficeXmlNode properties = PropertiesElement;
                    OfficeXmlNode before = element.NextSibling;
                    OfficeXmlNode lineBreak = parent.InsertBefore(parent.CreateElement(OfficeXmlNamespaces.DrawingML, "br"), before);
                    PowerPointTextContent.ApplyProperties(lineBreak, properties);
                    continuation.Add(lineBreak);
                    continuation.AddRange(PowerPointTextContent.InsertInlineContent(parent, rest, properties, before));
                    _lineContinuation = continuation;
                }
                MarkEdited();
            }
        }

        /// <summary>
        /// Changes the stored run text casing while preserving run formatting.
        /// </summary>
        public PowerPointTextRun TransformTextCase(OfficeTextCase textCase, CultureInfo? culture = null) {
            Text = OfficeTextCaseTransformer.Apply(Text, textCase, culture);
            return this;
        }

        private bool ReadFlag(string attribute) => PropertiesElement.BooleanAttribute(attribute) == true;

        // Sets an attribute of the run's a:rPr, adding the a:rPr first in the run when missing; null removes it.
        private void SetProperty(string attribute, string? value) {
            EnsurePropertiesElement().SetAttribute(attribute, value);
            SyncLineContinuation();
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is bold.
        /// </summary>
        public bool Bold {
            get => ReadFlag("b");
            set => SetProperty("b", value ? "1" : "0");
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is italic.
        /// </summary>
        public bool Italic {
            get => ReadFlag("i");
            set => SetProperty("i", value ? "1" : "0");
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is underlined.
        /// </summary>
        public bool Underline {
            get => UnderlineStyle is { } style && style != PowerPointUnderlineStyle.None;
            set {
                UnderlineStyle = value ? PowerPointUnderlineStyle.Single : PowerPointUnderlineStyle.None;
            }
        }

        /// <summary>
        /// Gets or sets the native DrawingML underline variant.
        /// </summary>
        public PowerPointUnderlineStyle? UnderlineStyle {
            get => PropertiesElement.Attribute("u") is { } token && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out PowerPointUnderlineStyle style)
                ? style : null;
            set => SetProperty("u", value?.ToOpenXmlToken());
        }

        /// <summary>
        /// Gets or sets a value indicating whether the run is strikethrough.
        /// </summary>
        public bool Strikethrough {
            get => StrikeStyle is { } style && style != PowerPointStrikeStyle.None;
            set {
                StrikeStyle = value ? PowerPointStrikeStyle.Single : PowerPointStrikeStyle.None;
            }
        }

        /// <summary>
        /// Gets or sets the native DrawingML strike-through variant.
        /// </summary>
        public PowerPointStrikeStyle? StrikeStyle {
            get => PropertiesElement.Attribute("strike") switch {
                null => null,
                "noStrike" => PowerPointStrikeStyle.None,
                "sngStrike" => PowerPointStrikeStyle.Single,
                "dblStrike" => PowerPointStrikeStyle.Double,
                string other => throw new InvalidOperationException($"Unsupported DrawingML strike value '{other}'.")
            };
            set => SetProperty("strike", value switch {
                null => null,
                PowerPointStrikeStyle.None => "noStrike",
                PowerPointStrikeStyle.Single => "sngStrike",
                PowerPointStrikeStyle.Double => "dblStrike",
                _ => throw new ArgumentOutOfRangeException(nameof(value))
            });
        }

        /// <summary>
        /// Gets or sets native DrawingML capitalization without changing the stored characters.
        /// </summary>
        public PowerPointCapitalization? Capitalization {
            get => PropertiesElement.Attribute("cap") switch {
                null => null,
                "none" => PowerPointCapitalization.None,
                "small" => PowerPointCapitalization.SmallCaps,
                "all" => PowerPointCapitalization.AllCaps,
                string other => throw new InvalidOperationException($"Unsupported DrawingML capitalization value '{other}'.")
            };
            set => SetProperty("cap", value switch {
                null => null,
                PowerPointCapitalization.None => "none",
                PowerPointCapitalization.SmallCaps => "small",
                PowerPointCapitalization.AllCaps => "all",
                _ => throw new ArgumentOutOfRangeException(nameof(value))
            });
        }

        /// <summary>
        /// Gets or sets the DrawingML baseline shift in percent, from -100 through 100.
        /// Positive values create superscript and negative values create subscript.
        /// </summary>
        public double? BaselinePercent {
            get => PropertiesElement.Int32Attribute("baseline") is int value ? value / 1000D : null;
            set {
                if (value.HasValue && (double.IsNaN(value.Value) || double.IsInfinity(value.Value)
                    || value.Value < -100D || value.Value > 100D)) {
                    throw new ArgumentOutOfRangeException(nameof(value), "Baseline percent must be between -100 and 100.");
                }
                int? baseline = value.HasValue ? checked((int)Math.Round(value.Value * 1000D, MidpointRounding.AwayFromZero)) : null;
                SetProperty("baseline", baseline?.ToString(CultureInfo.InvariantCulture));
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
            BaselinePercent = 0D;
            return this;
        }

        /// <summary>
        /// Gets or sets the physical font size, preserving DrawingML hundredth-point precision.
        /// Null removes the override; sizes must be from 1 through 4000 points.
        /// </summary>
        public global::OfficeIMO.OfficeLength? FontSize {
            get => PropertiesElement.Int32Attribute("sz") is int size ? global::OfficeIMO.OfficeLength.Points(size / 100D) : (global::OfficeIMO.OfficeLength?)null;
            set {
                int? drawingSize = PowerPointTextDefaults.ToDrawingFontSize(value?.ToPoints(), nameof(value));
                SetProperty("sz", drawingSize?.ToString(CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Gets or sets the font family (Latin).
        /// </summary>
        public string? FontFamily {
            get => PropertiesElement.Child(OfficeXmlNamespaces.DrawingML, "latin").Attribute("typeface");
            set {
                PowerPointTextPropertyFormatting.SetFontFamily(EnsurePropertiesElement(), value);
                SyncLineContinuation();
            }
        }

        /// <summary>
        /// Gets or sets the local text color, retaining native theme slots. Alpha below 255 is written as native DrawingML opacity.
        /// Explicit color also applies to hyperlinks in readers supporting the native text-color extension.
        /// </summary>
        public global::OfficeIMO.OfficeColor? Color {
            get => PowerPointDrawingColor.Read(PropertiesElement.Child(OfficeXmlNamespaces.DrawingML, "solidFill"));
            set {
                // The theme can open the SDK, so it is read before anything is edited.
                OfficeTheme? theme = PowerPointDrawingColor.NeedsTheme(value) ? PowerPointDrawingColor.ThemeFor(OwnerPartRef, Slide) : null;
                PowerPointTextPropertyFormatting.SetColor(EnsurePropertiesElement(), value, theme);
                SyncLineContinuation();
            }
        }

        /// <summary>
        /// Gets or sets the run background (DrawingML highlight) color. Alpha below 255 is written as native opacity.
        /// </summary>
        public global::OfficeIMO.OfficeColor? BackgroundColor {
            get => PowerPointDrawingColor.Read(PropertiesElement.Child(OfficeXmlNamespaces.DrawingML, "highlight"));
            set {
                OfficeTheme? theme = PowerPointDrawingColor.NeedsTheme(value) ? PowerPointDrawingColor.ThemeFor(OwnerPartRef, Slide) : null;
                PowerPointTextPropertyFormatting.SetBackgroundColor(EnsurePropertiesElement(), value, theme);
                SyncLineContinuation();
            }
        }

        private OfficeXmlNode ClickHyperlinkElement => PropertiesElement.Child(OfficeXmlNamespaces.DrawingML, "hlinkClick");

        /// <summary>
        /// Gets or sets the hyperlink target for this run. Internal slide links are returned as
        /// stable Markdown-compatible fragments such as <c>#slide-2</c>.
        /// </summary>
        public Uri? Hyperlink {
            get {
                PowerPointPartRef owner = OwnerPartRef;
                return owner.IsNone ? null : PowerPointHyperlinkResolver.Resolve(owner, SlidePartRef, ClickHyperlinkElement);
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
            ArgumentNullException.ThrowIfNull(url);

            SetHyperlink(new Uri(url, UriKind.RelativeOrAbsolute), tooltip);
        }

        /// <summary>
        /// Sets a hyperlink for this run.
        /// </summary>
        public void SetHyperlink(Uri uri, string? tooltip = null) {
            ArgumentNullException.ThrowIfNull(uri);
            PowerPointPartRef slidePart = SlidePartRef;
            if (!uri.IsAbsoluteUri && !slidePart.IsNone && slidePart.Document is { } document
                && PowerPointHyperlinkResolver.ResolveSlideFragment(document, uri) is { IsNone: false } target) {
                SetInternalHyperlink(target, tooltip);
                return;
            }
            PowerPointPartRef owner = OwnerPartRef;
            if (owner.IsNone) {
                throw new InvalidOperationException("Hyperlinks require an owning presentation part.");
            }

            string relationshipId = owner.AddExternalHyperlink(uri);
            ReplaceClickHyperlink(relationshipId, action: null, tooltip);
        }

        /// <summary>
        /// Sets an internal hyperlink from this run to another slide in the same presentation.
        /// </summary>
        public void SetHyperlink(PowerPointSlide targetSlide,
            string? tooltip = null) {
            ArgumentNullException.ThrowIfNull(targetSlide);
            SetInternalHyperlink(targetSlide.PartRef, tooltip);
        }

        private void SetInternalHyperlink(PowerPointPartRef targetSlidePart, string? tooltip) {
            PowerPointPartRef slidePart = SlidePartRef;
            if (slidePart.IsNone) {
                throw new InvalidOperationException(
                    "Hyperlinks require a slide context.");
            }
            PowerPointPartRef owner = !OwnerPartRef.IsNone ? OwnerPartRef : slidePart;
            if (slidePart.Document is null || !ReferenceEquals(slidePart.Document, targetSlidePart.Document)) {
                throw new ArgumentException(
                    "The hyperlink target must belong to the same presentation.",
                    "targetSlide");
            }

            PowerPointPartKind ownerKind = owner.Kind;
            string relationshipId = ownerKind is PowerPointPartKind.SlideLayout or PowerPointPartKind.SlideMaster
                || ownerKind == PowerPointPartKind.NotesSlide && !slidePart.Equals(targetSlidePart)
                ? owner.GetOrAddRelativeHyperlink(targetSlidePart)
                : owner.GetOrAddSlideRelationship(targetSlidePart);
            ReplaceClickHyperlink(relationshipId, PowerPointHyperlinkResolver.SlideJumpAction, tooltip);
        }

        /// <summary>
        /// Removes any hyperlink from this run.
        /// </summary>
        public void ClearHyperlink() {
            if (PropertiesElement.IsNone) return;
            ReplaceClickHyperlink(relationshipId: null, action: null, tooltip: null);
        }

        // Replaces the run's click hyperlinks with one to the relationship (none when null), keeping the color policy
        // and sound of the previous one, and removes relationships nothing uses any more.
        private void ReplaceClickHyperlink(string? relationshipId, string? action, string? tooltip) {
            OfficeXmlNode properties = relationshipId == null ? PropertiesElement : EnsurePropertiesElement();
            List<OfficeXmlNode> previous = properties.Children(OfficeXmlNamespaces.DrawingML, "hlinkClick").ToList();
            bool? colorChoice = null;
            OfficeXmlNode preservedSound = default;
            bool? preservedEndSound = null;
            var relationshipIds = new List<string>();
            var soundRelationshipIds = new List<string>();
            foreach (OfficeXmlNode link in previous) {
                colorChoice ??= PowerPointTextHyperlinkColor.ReadChoice(link);
                OfficeXmlNode sound = link.Child(OfficeXmlNamespaces.DrawingML, "snd");
                if (preservedSound.IsNone && !sound.IsNone) preservedSound = sound.Clone();
                preservedEndSound ??= link.BooleanAttribute("endSnd");
                if (link.Attribute(OfficeXmlNamespaces.Relationships, "id") is { Length: > 0 } id && !relationshipIds.Contains(id)) relationshipIds.Add(id);
                foreach (OfficeXmlNode linkSound in link.Children(OfficeXmlNamespaces.DrawingML, "snd")) {
                    if (linkSound.Attribute(OfficeXmlNamespaces.Relationships, "embed") is { Length: > 0 } embed && !soundRelationshipIds.Contains(embed)) soundRelationshipIds.Add(embed);
                }
            }
            foreach (OfficeXmlNode link in previous) link.Remove();
            if (relationshipId != null) {
                OfficeXmlNode replacement = properties.CreateElement(OfficeXmlNamespaces.DrawingML, "hlinkClick");
                replacement.SetAttribute(OfficeXmlNamespaces.Relationships, "id", relationshipId);
                if (action != null) replacement.SetAttribute("action", action);
                if (!string.IsNullOrWhiteSpace(tooltip)) replacement.SetAttribute("tooltip", tooltip);
                if (preservedEndSound.HasValue) replacement.SetBooleanAttribute("endSnd", preservedEndSound.Value);
                replacement = properties.InsertOrdered(replacement, ATextCharacterPropertiesView.Order);
                if (!preservedSound.IsNone) replacement.InsertBefore(preservedSound, replacement.FirstChild);
                if (colorChoice.HasValue || !properties.Child(OfficeXmlNamespaces.DrawingML, "solidFill").IsNone) {
                    PowerPointTextHyperlinkColor.SetChoice(replacement, colorChoice ?? true);
                }
            }
            MarkEdited();
            SyncLineContinuation();
            PowerPointPartRef owner = OwnerPartRef;
            if (owner.IsNone) return;
            foreach (string id in relationshipIds) owner.RemoveRelationshipIfUnused(id, SlidePartRef);
            foreach (string id in soundRelationshipIds) owner.RemoveRelationshipIfUnused(id, default);
        }

        /// <summary>Removes a content subtree and the relationships no surviving content uses.</summary>
        internal static void DiscardContent(OfficeXmlNode content, PowerPointPartRef owner, PowerPointPartRef owningSlide = default) {
            var ids = new List<string>();
            CollectRelationshipIds(content, ids);
            if (!content.Parent.IsNone) content.Remove();
            if (owner.IsNone) return;
            owner.MarkEdited();
            // A notes slide keeps its relationship to the slide it belongs to.
            PowerPointPartRef slide = owner.Kind == PowerPointPartKind.Slide ? owner : owningSlide;
            foreach (string id in ids) owner.RemoveRelationshipIfUnused(id, slide);
        }

        // The r:id of every hyperlink and the r:embed of every sound below the element, once each.
        private static void CollectRelationshipIds(OfficeXmlNode element, List<string> ids) {
            foreach (OfficeXmlNode child in element.Children()) {
                if (child.NamespaceId == OfficeXmlNamespaces.DrawingML) {
                    string? id = child.LocalName switch {
                        "hlinkClick" or "hlinkMouseOver" or "hlinkHover" => child.Attribute(OfficeXmlNamespaces.Relationships, "id"),
                        "snd" => child.Attribute(OfficeXmlNamespaces.Relationships, "embed"),
                        _ => null
                    };
                    if (!string.IsNullOrEmpty(id) && !ids.Contains(id!)) ids.Add(id!);
                }
                CollectRelationshipIds(child, ids);
            }
        }

        private void SyncLineContinuation() {
            if (_lineContinuation == null) return;
            OfficeXmlNode properties = PropertiesElement;
            foreach (OfficeXmlNode line in _lineContinuation) PowerPointTextContent.ApplyProperties(line, properties);
        }
    }
}
