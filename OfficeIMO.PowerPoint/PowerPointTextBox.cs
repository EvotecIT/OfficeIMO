using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a textbox shape.
    /// </summary>
    public partial class PowerPointTextBox : PowerPointShape {
        // The children of p:txBody and a:bodyPr, in sequence order.
        private static readonly OfficeXmlChildOrder BodyOrder = new(
            new[] { OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML }, new[] { "bodyPr", "lstStyle", "p" });
        private static readonly OfficeXmlChildOrder BodyPropertiesOrder = new(
            new[] { OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML },
            new[] { "prstTxWarp", "noAutofit", "normAutofit", "spAutoFit", "scene3d", "sp3d", "flatTx", "extLst" },
            new[] { 0, 1, 1, 1, 2, 3, 3, 4 });
        // The children of p:nvPr, in sequence order.
        private static readonly OfficeXmlChildOrder ApplicationPropertiesOrder = new(
            new[] { OfficeXmlNamespaces.PresentationML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.PresentationML, OfficeXmlNamespaces.PresentationML },
            new[] { "ph", "audioCd", "wavAudioFile", "audioFile", "videoFile", "quickTimeFile", "extLst", "custDataLst", "extLst" },
            new[] { 0, 1, 1, 1, 1, 1, 1, 2, 3 });

        internal PowerPointTextBox(PowerPointSlide slide, OfficeXmlTree tree, int node) : base(slide, tree, node) {
        }

        internal PowerPointTextBox(OpenXml.PowerPointPartRef part, OfficeXmlTree tree, int node) : base(part, tree, node) {
        }

        private OfficeXmlNode ShapeElement => XmlElement;

        private OfficeXmlNode BodyElement => ShapeElement.Child(OfficeXmlNamespaces.PresentationML, "txBody");

        private OfficeXmlNode BodyPropertiesElement => BodyElement.Child(OfficeXmlNamespaces.DrawingML, "bodyPr");

        private OfficeXmlNode PlaceholderElement =>
            ShapeElement.Child(OfficeXmlNamespaces.PresentationML, "nvSpPr").Child(OfficeXmlNamespaces.PresentationML, "nvPr").Child(OfficeXmlNamespaces.PresentationML, "ph");

        private void MarkEdited() => OwnerPartRef.MarkEdited();

        /// <summary>
        /// Gets the preset geometry used by this text-bearing shape.
        /// </summary>
        public OfficePresetShapeType? ShapeType =>
            ShapeElement.Child(OfficeXmlNamespaces.PresentationML, "spPr").Child(OfficeXmlNamespaces.DrawingML, "prstGeom").Attribute("prst") is { } token
            && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out OfficePresetShapeType type) ? type : null;

        /// <summary>
        /// Gets whether this shape represents a conventional text box or placeholder rather than a text-bearing preset shape.
        /// </summary>
        public bool UsesTextBoxGeometry {
            get {
                OfficeXmlNode nonVisual = ShapeElement.Child(OfficeXmlNamespaces.PresentationML, "nvSpPr");
                return nonVisual.Child(OfficeXmlNamespaces.PresentationML, "cNvSpPr").BooleanAttribute("txBox") == true
                    || !nonVisual.Child(OfficeXmlNamespaces.PresentationML, "nvPr").Child(OfficeXmlNamespaces.PresentationML, "ph").IsNone;
            }
        }

        /// <summary>
        ///     Paragraphs contained in the textbox.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> Paragraphs {
            get {
                var paragraphs = new List<PowerPointParagraph>();
                foreach (OfficeXmlNode paragraph in EnsureBody().Children(OfficeXmlNamespaces.DrawingML, "p")) paragraphs.Add(ParagraphFor(paragraph));
                return paragraphs;
            }
        }

        // A paragraph wrapper bound the same way the shape is; a node taken from the package tree binds by its place
        // there, even when the SDK opened after it was taken.
        private PowerPointParagraph ParagraphFor(OfficeXmlNode paragraph) =>
            paragraph.Tree != null && OwnerSlide != null && ReferenceEquals(paragraph.Tree, PackageTree)
                ? new PowerPointParagraph(OwnerSlide, paragraph.Tree, paragraph.Node)
                : new PowerPointParagraph(paragraph, SlidePartRef, OwnerPartRef);

        /// <summary>
        ///     Adds a paragraph to the textbox.
        /// </summary>
        public PowerPointParagraph AddParagraph(string text = "", Action<PowerPointParagraph>? configure = null,
            Action<PowerPointTextRun>? run = null) {
            OfficeXmlNode body = EnsureBody();
            PowerPointParagraph wrapper = AppendParagraph(body, text ?? string.Empty, body.Child(OfficeXmlNamespaces.DrawingML, "p"), out PowerPointTextRun authoredRun);
            configure?.Invoke(wrapper);
            // Configuring the paragraph can replace its runs; the callback then formats a run that is still in it.
            if (run != null) run(authoredRun.BelongsTo(wrapper) ? authoredRun : wrapper.AddText(text ?? string.Empty));
            return wrapper;
        }

        /// <summary>
        ///     Adds multiple paragraphs to the textbox.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> AddParagraphs(IEnumerable<string> paragraphs,
            Action<PowerPointParagraph>? configure = null) {
            ArgumentNullException.ThrowIfNull(paragraphs);

            var results = new List<PowerPointParagraph>();
            foreach (string paragraphText in paragraphs) {
                PowerPointParagraph paragraph = AddParagraph(paragraphText ?? string.Empty);
                configure?.Invoke(paragraph);
                results.Add(paragraph);
            }

            return results;
        }

        /// <summary>
        ///     Replaces all paragraphs with the provided content.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> SetParagraphs(IEnumerable<string> paragraphs,
            Action<PowerPointParagraph>? configure = null) {
            ArgumentNullException.ThrowIfNull(paragraphs);

            return ReplaceParagraphs(paragraphs, (paragraphText, templateParagraph) => {
                PowerPointParagraph paragraph = AppendParagraph(EnsureBody(), paragraphText ?? string.Empty, templateParagraph, out _);
                configure?.Invoke(paragraph);
                return paragraph;
            });
        }

        /// <summary>
        ///     Applies a text style to the textbox.
        /// </summary>
        public PowerPointTextBox ApplyTextStyle(PowerPointTextStyle style, bool applyToRuns = false) {
            var paragraphs = GetParagraphsForStyling();
            foreach (PowerPointParagraph paragraph in paragraphs) {
                if (applyToRuns) {
                    IReadOnlyList<PowerPointTextRun> runs = paragraph.Runs;
                    if (runs.Count == 0) {
                        runs = new List<PowerPointTextRun> { paragraph.AddText(string.Empty) };
                    }
                    foreach (PowerPointTextRun run in runs) {
                        style.Apply(run);
                    }
                } else {
                    style.Apply(paragraph);
                }
            }

            return this;
        }

        /// <summary>
        ///     Applies a paragraph style to the textbox.
        /// </summary>
        public PowerPointTextBox ApplyParagraphStyle(PowerPointParagraphStyle style) {
            var paragraphs = GetParagraphsForStyling();
            foreach (PowerPointParagraph paragraph in paragraphs) {
                style.Apply(paragraph);
            }

            return this;
        }

        /// <summary>
        ///     Applies automatic spacing defaults to the textbox paragraphs.
        /// </summary>
        public PowerPointTextBox ApplyAutoSpacing(double lineSpacingMultiplier = 1.2,
            global::OfficeIMO.OfficeLength? spaceBefore = null, global::OfficeIMO.OfficeLength? spaceAfter = null) {
            PowerPointParagraphStyle style = new(lineSpacingMultiplier: lineSpacingMultiplier,
                spaceBefore: spaceBefore,
                spaceAfter: spaceAfter);
            return ApplyParagraphStyle(style);
        }

        /// <summary>
        ///     Replaces literal text in the textbox paragraphs.
        /// </summary>
        /// <param name="text">Text to find.</param>
        /// <param name="replacement">Replacement text; <c>null</c> removes each match.</param>
        /// <param name="options">Comparison settings. <c>null</c> matches ordinally, ignoring case, inside or across words.</param>
        /// <returns>The number of replaced matches.</returns>
        /// <remarks>A match can span adjacent runs within one paragraph. Replacement text takes the formatting of the run where the match starts.</remarks>
        public int ReplaceAll(string text, string replacement, OfficeTextSearchOptions? options = null) {
            OfficeTextSearch.ValidateSearchText(text, nameof(text));
            int count = 0;
            foreach (PowerPointParagraph paragraph in Paragraphs) {
                count += paragraph.ReplaceAll(text, replacement, options);
            }
            return count;
        }

        /// <summary>
        ///     Removes all paragraphs from the textbox.
        /// </summary>
        public void Clear() {
            OfficeXmlNode body = EnsureBody();
            List<string> soundIds = SoundIds(body);
            OfficeXmlNode template = body.Child(OfficeXmlNamespaces.DrawingML, "p");
            OfficeXmlNode empty = CreateEmptyParagraph(body, template);
            body.RemoveChildren(OfficeXmlNamespaces.DrawingML, "p");
            body.Append(empty);
            MarkEdited();
            RemoveUnusedSounds(soundIds);
        }

        /// <summary>
        ///     Text contained in the textbox.
        /// </summary>
        public string Text {
            get => PowerPointParagraph.ReadBodyText(EnsureBody());
            set {
                OfficeXmlNode body = EnsureBody();
                List<string> soundIds = SoundIds(body);
                OfficeXmlNode template = body.Child(OfficeXmlNamespaces.DrawingML, "p");
                // Each line becomes a paragraph with copies of the first paragraph's properties, its first run's
                // formatting and its end properties, taken before the old paragraphs go.
                OfficeXmlNode paragraphProperties = template.Child(OfficeXmlNamespaces.DrawingML, "pPr");
                paragraphProperties = paragraphProperties.IsNone ? default : paragraphProperties.Clone();
                OfficeXmlNode endProperties = template.Child(OfficeXmlNamespaces.DrawingML, "endParaRPr");
                endProperties = endProperties.IsNone ? default : CloneWithoutSounds(endProperties);
                OfficeXmlNode runProperties = FirstRunProperties(template);
                runProperties = runProperties.IsNone ? default : CloneWithoutSounds(runProperties);

                body.RemoveChildren(OfficeXmlNamespaces.DrawingML, "p");
                foreach (string line in PowerPointTextContent.SplitParagraphs(value ?? string.Empty)) {
                    OfficeXmlNode paragraph = body.Append(body.CreateElement(OfficeXmlNamespaces.DrawingML, "p"));
                    if (!paragraphProperties.IsNone) paragraph.Append(paragraphProperties.Clone());
                    PowerPointTextContent.InsertInlineContent(paragraph, line, runProperties, default);
                    if (!endProperties.IsNone) paragraph.Append(endProperties.Clone());
                }
                MarkEdited();
                RemoveUnusedSounds(soundIds);
            }
        }

        /// <summary>
        ///     Returns true if the textbox is tied to a slide placeholder.
        /// </summary>
        public bool IsPlaceholder => !PlaceholderElement.IsNone;

        /// <summary>
        ///     Gets or sets the placeholder type for this textbox.
        /// </summary>
        public PowerPointPlaceholderType? PlaceholderType {
            get => PlaceholderElement.Attribute("type") is { } token && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out PowerPointPlaceholderType type) ? type : null;
            set => SetPlaceholderAttribute("type", value?.ToOpenXmlToken());
        }

        /// <summary>
        ///     Gets or sets the placeholder index for this textbox.
        /// </summary>
        public uint? PlaceholderIndex {
            get => PlaceholderElement.UInt32Attribute("idx");
            set => SetPlaceholderAttribute("idx", OfficeXmlViewValues.Format(value));
        }

        /// <summary>
        ///     Gets or sets the preferred placeholder size.
        /// </summary>
        public PowerPointPlaceholderSize? PlaceholderSize {
            get => PlaceholderElement.Attribute("sz") is { } token && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out PowerPointPlaceholderSize size) ? size : null;
            set => SetPlaceholderAttribute("sz", value?.ToOpenXmlToken());
        }

        /// <summary>
        ///     Gets or sets the placeholder orientation.
        /// </summary>
        public PowerPointPlaceholderDirection? PlaceholderOrientation {
            get => PlaceholderElement.Attribute("orient") is { } token && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out PowerPointPlaceholderDirection direction) ? direction : null;
            set => SetPlaceholderAttribute("orient", value?.ToOpenXmlToken());
        }

        // Sets a p:ph attribute, adding the placeholder when needed; clearing the last attribute of an empty placeholder
        // removes it.
        private void SetPlaceholderAttribute(string attribute, string? value) {
            if (value == null) {
                OfficeXmlNode placeholder = PlaceholderElement;
                if (placeholder.IsNone) return;
                placeholder.SetAttribute(attribute, null);
                if (CanRemovePlaceholder(placeholder)) placeholder.Remove();
                MarkEdited();
                return;
            }
            EnsurePlaceholder().SetAttribute(attribute, value);
            MarkEdited();
        }

        /// <summary>
        ///     Gets or sets the left text margin as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? TextMarginLeft {
            get => ReadInset("lIns");
            set => SetBodyAttribute("lIns", OfficeXmlViewValues.Format(ToEmusInt(value?.ToPoints(), PowerPointUnits.EmusPerPoint)));
        }

        /// <summary>
        ///     Gets or sets the right text margin as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? TextMarginRight {
            get => ReadInset("rIns");
            set => SetBodyAttribute("rIns", OfficeXmlViewValues.Format(ToEmusInt(value?.ToPoints(), PowerPointUnits.EmusPerPoint)));
        }

        /// <summary>
        ///     Gets or sets the top text margin as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? TextMarginTop {
            get => ReadInset("tIns");
            set => SetBodyAttribute("tIns", OfficeXmlViewValues.Format(ToEmusInt(value?.ToPoints(), PowerPointUnits.EmusPerPoint)));
        }

        /// <summary>
        ///     Gets or sets the bottom text margin as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? TextMarginBottom {
            get => ReadInset("bIns");
            set => SetBodyAttribute("bIns", OfficeXmlViewValues.Format(ToEmusInt(value?.ToPoints(), PowerPointUnits.EmusPerPoint)));
        }

        private global::OfficeIMO.OfficeLength? ReadInset(string attribute) =>
            BodyPropertiesElement.Int32Attribute(attribute) is int value ? global::OfficeIMO.OfficeLength.Emus(value) : (global::OfficeIMO.OfficeLength?)null;

        private void SetBodyAttribute(string attribute, string? value) {
            EnsureBodyProperties().SetAttribute(attribute, value);
            MarkEdited();
        }

        /// <summary>
        ///     Sets all text margins as a physical length.
        /// </summary>
        public PowerPointTextBox SetTextMargins(global::OfficeIMO.OfficeLength left, global::OfficeIMO.OfficeLength top, global::OfficeIMO.OfficeLength right, global::OfficeIMO.OfficeLength bottom) {
            TextMarginLeft = left;
            TextMarginTop = top;
            TextMarginRight = right;
            TextMarginBottom = bottom;
            return this;
        }

        /// <summary>
        ///     Gets or sets the vertical anchoring of text inside the textbox.
        /// </summary>
        public PowerPointTextVerticalAlignment? TextVerticalAlignment {
            get => BodyPropertiesElement.Attribute("anchor") is { } token && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out PowerPointTextVerticalAlignment anchor) ? anchor : null;
            set => SetBodyAttribute("anchor", value?.ToOpenXmlToken());
        }

        /// <summary>
        ///     Gets or sets the text direction (horizontal/vertical).
        /// </summary>
        public PowerPointTextDirection? TextDirection {
            get => BodyPropertiesElement.Attribute("vert") is { } token && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out PowerPointTextDirection direction) ? direction : null;
            set => SetBodyAttribute("vert", value?.ToOpenXmlToken());
        }

        /// <summary>
        ///     Gets or sets the text auto-fit behavior.
        /// </summary>
        public PowerPointTextAutoFit? TextAutoFit {
            get {
                OfficeXmlNode body = BodyPropertiesElement;
                if (body.IsNone) return null;
                if (!body.Child(OfficeXmlNamespaces.DrawingML, "noAutofit").IsNone) return PowerPointTextAutoFit.None;
                if (!body.Child(OfficeXmlNamespaces.DrawingML, "normAutofit").IsNone) return PowerPointTextAutoFit.Normal;
                if (!body.Child(OfficeXmlNamespaces.DrawingML, "spAutoFit").IsNone) return PowerPointTextAutoFit.Shape;
                return null;
            }
            set {
                OfficeXmlNode body = EnsureBodyProperties();
                RemoveAutoFit(body, "noAutofit", "normAutofit", "spAutoFit");
                if (value != null) {
                    body.InsertOrdered(body.CreateElement(OfficeXmlNamespaces.DrawingML, value.Value switch {
                        PowerPointTextAutoFit.None => "noAutofit",
                        PowerPointTextAutoFit.Normal => "normAutofit",
                        _ => "spAutoFit"
                    }), BodyPropertiesOrder);
                }
                MarkEdited();
            }
        }

        private static void RemoveAutoFit(OfficeXmlNode body, params string[] names) {
            foreach (string name in names) body.RemoveChildren(OfficeXmlNamespaces.DrawingML, name);
        }

        /// <summary>
        ///     Gets or sets detailed auto-fit options (only applies to Normal auto-fit).
        /// </summary>
        public PowerPointTextAutoFitOptions? TextAutoFitOptions {
            get {
                OfficeXmlNode normal = BodyPropertiesElement.Child(OfficeXmlNamespaces.DrawingML, "normAutofit");
                return normal.IsNone ? null : PowerPointTextAutoFitOptions.FromOpenXmlValues(normal.Int32Attribute("fontScale"), normal.Int32Attribute("lnSpcReduction"));
            }
            set {
                if (value == null) {
                    OfficeXmlNode existing = BodyPropertiesElement;
                    if (!existing.IsNone) {
                        existing.RemoveChildren(OfficeXmlNamespaces.DrawingML, "normAutofit");
                        MarkEdited();
                    }
                    return;
                }

                OfficeXmlNode body = EnsureBodyProperties();
                RemoveAutoFit(body, "noAutofit", "spAutoFit");
                ApplyNormalAutoFitOptions(NormalAutoFit(body), value.Value);
                MarkEdited();
            }
        }

        private static OfficeXmlNode NormalAutoFit(OfficeXmlNode body) {
            OfficeXmlNode normal = body.Child(OfficeXmlNamespaces.DrawingML, "normAutofit");
            return normal.IsNone ? body.InsertOrdered(body.CreateElement(OfficeXmlNamespaces.DrawingML, "normAutofit"), BodyPropertiesOrder) : normal;
        }

        /// <summary>
        ///     Sets the auto-fit mode and optional Normal auto-fit options.
        /// </summary>
        public PowerPointTextBox SetTextAutoFit(PowerPointTextAutoFit fit, PowerPointTextAutoFitOptions? options = null) {
            TextAutoFit = fit;
            if (fit == PowerPointTextAutoFit.Normal && options != null) {
                ApplyNormalAutoFitOptions(NormalAutoFit(EnsureBodyProperties()), options.Value);
                MarkEdited();
            }
            return this;
        }

        // The a:rPr of the first run of the body's paragraphs, or none.
        private OfficeXmlNode FirstRunRunProperties {
            get {
                foreach (OfficeXmlNode paragraph in BodyElement.Children(OfficeXmlNamespaces.DrawingML, "p")) {
                    OfficeXmlNode run = paragraph.Child(OfficeXmlNamespaces.DrawingML, "r");
                    if (!run.IsNone) return run.Child(OfficeXmlNamespaces.DrawingML, "rPr");
                }
                return default;
            }
        }

        /// <summary>
        ///     Gets or sets a value indicating whether the text is bold.
        /// </summary>
        public bool Bold {
            get => FirstRunRunProperties.BooleanAttribute("b") == true;
            set => FormatRuns(properties => properties.SetAttribute("b", value ? "1" : "0"));
        }

        /// <summary>
        ///     Gets or sets a value indicating whether the text is italic.
        /// </summary>
        public bool Italic {
            get => FirstRunRunProperties.BooleanAttribute("i") == true;
            set => FormatRuns(properties => properties.SetAttribute("i", value ? "1" : "0"));
        }

        /// <summary>
        ///     Gets or sets the font size as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? FontSize {
            get => FirstRunRunProperties.Int32Attribute("sz") is int size ? global::OfficeIMO.OfficeLength.Points(size / 100D) : (global::OfficeIMO.OfficeLength?)null;
            set {
                int? drawingSize = PowerPointTextDefaults.ToDrawingFontSize(value?.ToPoints(), nameof(value));
                FormatRuns(properties => properties.SetAttribute("sz", OfficeXmlViewValues.Format(drawingSize)));
            }
        }

        /// <summary>
        ///     Gets or sets the font family.
        /// </summary>
        public string? FontFamily {
            get => FirstRunRunProperties.Child(OfficeXmlNamespaces.DrawingML, "latin").Attribute("typeface");
            set => FormatRuns(properties => PowerPointTextPropertyFormatting.SetFontFamily(properties, value));
        }

        /// <summary>
        ///     Gets the first run's RGB text color or sets the color of every run. Alpha below 255 is written as native opacity.
        /// </summary>
        public global::OfficeIMO.OfficeColor? Color {
            get => PowerPointDrawingColor.Read(FirstRunRunProperties.Child(OfficeXmlNamespaces.DrawingML, "solidFill"));
            set {
                // The theme can open the SDK, so it is read before anything is edited.
                OfficeTheme? theme = PowerPointDrawingColor.NeedsTheme(value) ? PowerPointDrawingColor.ThemeFor(OwnerPartRef, OwnerSlide) : null;
                FormatRuns(properties => PowerPointTextPropertyFormatting.SetColor(properties, value, theme));
            }
        }

        // Applies a change to the properties of every run and line break of the text body.
        private void FormatRuns(Action<OfficeXmlNode> change) {
            PowerPointTextContent.FormatRuns(EnsureBody(), change);
            MarkEdited();
        }

        private static void ApplyNormalAutoFitOptions(OfficeXmlNode normal, PowerPointTextAutoFitOptions options) {
            normal.SetInt32Attribute("fontScale", options.FontScaleValue);
            normal.SetInt32Attribute("lnSpcReduction", options.LineSpaceReductionValue);
        }

        // The text body, added with empty body properties and list style when the shape has none.
        private OfficeXmlNode EnsureBody() {
            OfficeXmlNode shape = ShapeElement;
            OfficeXmlNode body = shape.Child(OfficeXmlNamespaces.PresentationML, "txBody");
            if (!body.IsNone) return body;
            OfficeXmlNode created = shape.CreateElement(OfficeXmlNamespaces.PresentationML, "txBody");
            created.Append(created.CreateElement(OfficeXmlNamespaces.DrawingML, "bodyPr"));
            created.Append(created.CreateElement(OfficeXmlNamespaces.DrawingML, "lstStyle"));
            body = shape.InsertOrdered(created, PShapeView.Order);
            MarkEdited();
            return body;
        }

        private List<PowerPointParagraph> GetParagraphsForStyling() {
            var paragraphs = Paragraphs.ToList();
            if (paragraphs.Count == 0) {
                paragraphs.Add(AddParagraph());
            }
            return paragraphs;
        }

        private IReadOnlyList<PowerPointParagraph> ReplaceParagraphs<T>(IEnumerable<T> items,
            Func<T, OfficeXmlNode, PowerPointParagraph> addParagraph) {
            // The items are read before anything is removed: a source that opens the SDK while it is read then does so
            // before the body below is taken.
            List<T> list = items as List<T> ?? items.ToList();
            OfficeXmlNode body = EnsureBody();
            List<string> soundIds = SoundIds(body);
            OfficeXmlNode template = body.Child(OfficeXmlNamespaces.DrawingML, "p");
            // The template leaves the body with the other paragraphs; the new ones copy from it.
            if (!template.IsNone) template.Remove();
            body.RemoveChildren(OfficeXmlNamespaces.DrawingML, "p");

            var results = new List<PowerPointParagraph>();
            foreach (T item in list) {
                results.Add(addParagraph(item, template));
            }

            if (results.Count == 0) {
                body.Append(CreateEmptyParagraph(body, template));
            }
            MarkEdited();
            RemoveUnusedSounds(soundIds);

            return results;
        }

        private PowerPointParagraph AppendParagraph(OfficeXmlNode body, string text, OfficeXmlNode template, out PowerPointTextRun authoredRun) {
            OfficeXmlNode paragraph = body.CreateElement(OfficeXmlNamespaces.DrawingML, "p");
            OfficeXmlNode paragraphProperties = template.Child(OfficeXmlNamespaces.DrawingML, "pPr");
            if (!paragraphProperties.IsNone) paragraph.Append(paragraphProperties.Clone());
            OfficeXmlNode endProperties = template.Child(OfficeXmlNamespaces.DrawingML, "endParaRPr");
            if (!endProperties.IsNone) paragraph.Append(CloneWithoutSounds(endProperties));
            OfficeXmlNode runProperties = FirstRunProperties(template);

            paragraph = body.Append(paragraph);
            MarkEdited();
            PowerPointParagraph wrapper = ParagraphFor(paragraph);
            authoredRun = wrapper.InsertRun(text, runProperties.IsNone ? default : CloneWithoutSounds(runProperties));
            return wrapper;
        }

        // A detached empty paragraph after the template: its properties, one empty run with its first run's formatting,
        // and its end properties.
        internal static OfficeXmlNode CreateEmptyParagraph(OfficeXmlNode context, OfficeXmlNode template) {
            OfficeXmlNode paragraph = context.CreateElement(OfficeXmlNamespaces.DrawingML, "p");
            OfficeXmlNode paragraphProperties = template.Child(OfficeXmlNamespaces.DrawingML, "pPr");
            if (!paragraphProperties.IsNone) paragraph.Append(paragraphProperties.Clone());
            OfficeXmlNode run = paragraph.Append(paragraph.CreateElement(OfficeXmlNamespaces.DrawingML, "r"));
            OfficeXmlNode runProperties = FirstRunProperties(template);
            if (!runProperties.IsNone) run.Append(CloneWithoutSounds(runProperties));
            run.Append(run.CreateElement(OfficeXmlNamespaces.DrawingML, "t"));
            OfficeXmlNode endProperties = template.Child(OfficeXmlNamespaces.DrawingML, "endParaRPr");
            if (!endProperties.IsNone) paragraph.Append(CloneWithoutSounds(endProperties));
            return paragraph;
        }

        // The a:rPr of the first run of a paragraph that has one, or none.
        internal static OfficeXmlNode FirstRunProperties(OfficeXmlNode paragraph) {
            foreach (OfficeXmlNode run in paragraph.Children(OfficeXmlNamespaces.DrawingML, "r")) {
                OfficeXmlNode properties = run.Child(OfficeXmlNamespaces.DrawingML, "rPr");
                if (!properties.IsNone) return properties;
            }
            return default;
        }

        // A copy of run or end properties whose hyperlinks no longer play the sounds the original text played.
        internal static OfficeXmlNode CloneWithoutSounds(OfficeXmlNode properties) {
            OfficeXmlNode copy = properties.Clone();
            foreach (OfficeXmlNode child in copy.Children().ToList()) {
                if (child.Is(OfficeXmlNamespaces.DrawingML, "hlinkClick") || child.Is(OfficeXmlNamespaces.DrawingML, "hlinkMouseOver")) {
                    child.RemoveChildren(OfficeXmlNamespaces.DrawingML, "snd");
                }
            }
            return copy;
        }

        internal static List<string> SoundIds(OfficeXmlNode element) {
            var ids = new List<string>();
            CollectSoundIds(element, ids);
            return ids;
        }

        private static void CollectSoundIds(OfficeXmlNode element, List<string> ids) {
            foreach (OfficeXmlNode child in element.Children()) {
                if (child.Is(OfficeXmlNamespaces.DrawingML, "snd") && child.Attribute(OfficeXmlNamespaces.Relationships, "embed") is { Length: > 0 } id && !ids.Contains(id)) ids.Add(id);
                CollectSoundIds(child, ids);
            }
        }

        private void RemoveUnusedSounds(List<string> soundIds) {
            if (soundIds.Count == 0) return;
            PowerPointPartRef owner = OwnerPartRef;
            if (owner.IsNone) return;
            foreach (string id in soundIds) owner.RemoveRelationshipIfUnused(id, default);
        }

        private OfficeXmlNode EnsureBodyProperties() {
            OfficeXmlNode body = EnsureBody();
            OfficeXmlNode properties = body.Child(OfficeXmlNamespaces.DrawingML, "bodyPr");
            return properties.IsNone ? body.InsertBefore(body.CreateElement(OfficeXmlNamespaces.DrawingML, "bodyPr"), body.FirstChild) : properties;
        }

        private OfficeXmlNode EnsurePlaceholder() {
            OfficeXmlNode shape = ShapeElement;
            OfficeXmlNode nonVisual = shape.GetOrAddChild(OfficeXmlNamespaces.PresentationML, "nvSpPr", PShapeView.Order);
            OfficeXmlNode application = nonVisual.GetOrAddChild(OfficeXmlNamespaces.PresentationML, "nvPr", PShapeNonVisualView.Order);
            return application.GetOrAddChild(OfficeXmlNamespaces.PresentationML, "ph", ApplicationPropertiesOrder);
        }

        private static bool CanRemovePlaceholder(OfficeXmlNode placeholder) =>
            placeholder.Attribute("type") == null && placeholder.Attribute("idx") == null
            && placeholder.Attribute("sz") == null && placeholder.Attribute("orient") == null
            && placeholder.Attribute("hasCustomPrompt") == null && placeholder.FirstChild.IsNone;

        private static int? ToEmusInt(double? value, double emusPerUnit) {
            return value != null ? (int)Math.Round(value.Value * emusPerUnit) : null;
        }

        private static double? FromEmusInt(int? emus, double emusPerUnit) {
            return emus != null ? emus.Value / emusPerUnit : null;
        }
    }
}
