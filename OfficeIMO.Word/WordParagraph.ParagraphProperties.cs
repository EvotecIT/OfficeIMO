using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    // Provides accessors for paragraph formatting.
    public partial class WordParagraph {
        /// <summary>
        /// Provides access to the borders applied to this paragraph.
        /// </summary>
        public WordParagraphBorders Borders {
            get {
                return new WordParagraphBorders(_document, this);
            }
        }

        /// <summary>
        /// Gets the identifier of the paragraph style, if any.
        /// </summary>
        public string? StyleId {
            get {
                if (_paragraphProperties != null && _paragraphProperties.ParagraphStyleId != null) {
                    return _paragraphProperties.ParagraphStyleId.Val;
                }
                return null;
            }
        }

        /// <summary>
        /// Alignment aka Paragraph Alignment. This element specifies the paragraph alignment which shall be applied to text in this paragraph.
        /// If this element is omitted on a given paragraph, its value is determined by the setting previously set at any level of the style hierarchy (i.e.that previous setting remains unchanged). If this setting is never specified in the style hierarchy, then no alignment is applied to the paragraph.
        /// Setting this property to <see langword="null"/> removes direct alignment and restores style inheritance.
        /// </summary>
        public WordParagraphAlignment? ParagraphAlignment {
            get {
                if (_paragraphProperties != null)
                    if (_paragraphProperties.Justification != null)
                        return _paragraphProperties.Justification.Val?.Value.ToOfficeEnum();
                return null;
            }
            set {
                if (value == null) {
                    _paragraphProperties?.RemoveAllChildren<Justification>();
                    return;
                }
                if (_paragraphProperties == null) {
                    _paragraph.ParagraphProperties = new ParagraphProperties();
                }
                _paragraph.ParagraphProperties!.Justification = new Justification {
                    Val = value.Value.ToOpenXml()
                };
            }
        }

        /// <summary>
        /// Text Alignment aka Vertical Character Alignment on Line. This element specifies the vertical alignment of all text on each line displayed within a paragraph. If the line height (before any added spacing) is larger than one or more characters on the line, all characters are aligned to each other as specified by this element.
        /// If this element is omitted on a given paragraph, its value is determined by the setting previously set at any level of the style hierarchy (i.e.that previous setting remains unchanged). If this setting is never specified in the style hierarchy, then the vertical alignment of all characters on the line shall be automatically determined by the consumer.
        /// Setting this property to <see langword="null"/> removes direct alignment and restores style inheritance.
        /// </summary>
        public WordVerticalCharacterAlignment? VerticalCharacterAlignmentOnLine {
            get {
                if (_paragraphProperties != null)
                    if (_paragraphProperties.TextAlignment != null)
                        return _paragraphProperties.TextAlignment.Val?.Value.ToOfficeEnum();
                return null;
            }
            set {
                if (value == null) {
                    _paragraphProperties?.RemoveAllChildren<TextAlignment>();
                    return;
                }
                if (_paragraphProperties == null) {
                    _paragraph.ParagraphProperties = new ParagraphProperties();
                }
                _paragraph.ParagraphProperties!.TextAlignment = new TextAlignment { Val = value.Value.ToOpenXml() };
            }
        }
        /// <summary>
        /// Gets or sets the indentation before the paragraph as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? IndentationBefore {
            get {
                if (_paragraphProperties != null && _paragraphProperties.Indentation != null) {
                    if (!string.IsNullOrEmpty(_paragraphProperties.Indentation.Left)) {
                        if (int.TryParse(_paragraphProperties.Indentation.Left, out var left)) {
                            return global::OfficeIMO.OfficeLength.Points((left) / 20D);
                        }
                    }
                }
                return null;
            }
            set {
                var wordNativeValue = checked((int?)(value?.ToTwips()));
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                var indentation = props.Indentation ?? new Indentation();
                indentation.Left = wordNativeValue.HasValue ? new StringValue(wordNativeValue.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)) : null;
                props.Indentation = indentation;
            }
        }
        /// <summary>
        /// Gets or sets the indentation after the paragraph as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? IndentationAfter {
            get {
                if (_paragraphProperties != null && _paragraphProperties.Indentation != null) {
                    if (!string.IsNullOrEmpty(_paragraphProperties.Indentation.Right)) {
                        if (int.TryParse(_paragraphProperties.Indentation.Right, out var right)) {
                            return global::OfficeIMO.OfficeLength.Points((right) / 20D);
                        }
                    }
                }
                return null;
            }
            set {
                var wordNativeValue = checked((int?)(value?.ToTwips()));
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                var indentation = props.Indentation ?? new Indentation();
                indentation.Right = wordNativeValue.HasValue ? new StringValue(wordNativeValue.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)) : null;
                props.Indentation = indentation;
            }
        }

        /// <summary>Gets or sets a directly authored page-break-before value; this does not resolve inherited paragraph styles.</summary>
        public bool PageBreakBefore {
            get {
                return ReadParagraphOnOff(_paragraphProperties?.PageBreakBefore);
            }
            set {
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                if (value == true) {
                    props.PageBreakBefore = new PageBreakBefore();
                } else {
                    props.PageBreakBefore = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets directly authored keep-with-next formatting without resolving paragraph styles.
        /// </summary>
        public bool KeepWithNext {
            get {
                return ReadParagraphOnOff(_paragraphProperties?.KeepNext);
            }
            set {
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                if (value) {
                    props.KeepNext ??= new KeepNext();
                    props.KeepNext.Val = null;
                } else {
                    props.KeepNext?.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets directly authored keep-lines-together formatting without resolving paragraph styles.
        /// </summary>
        public bool KeepLinesTogether {
            get {
                return ReadParagraphOnOff(_paragraphProperties?.KeepLines);
            }
            set {
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                if (value) {
                    props.KeepLines ??= new KeepLines();
                    props.KeepLines.Val = null;
                } else {
                    props.KeepLines?.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets directly authored widow and orphan control without resolving paragraph styles.
        /// </summary>
        public bool AvoidWidowAndOrphan {
            get {
                WidowControl? widowControl = _paragraphProperties?.WidowControl;
                return widowControl?.Val?.Value ?? widowControl is not null;
            }
            set {
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                if (value) {
                    props.WidowControl ??= new WidowControl();
                    props.WidowControl.Val = null;
                } else {
                    props.WidowControl = new WidowControl {
                        Val = false
                    };
                }
            }
        }
        /// <summary>
        /// Gets or sets the first line indentation as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? IndentationFirstLine {
            get {
                if (_paragraphProperties != null && _paragraphProperties.Indentation != null) {
                    if (!string.IsNullOrEmpty(_paragraphProperties.Indentation.FirstLine)) {
                        if (int.TryParse(_paragraphProperties.Indentation.FirstLine, out var firstLine)) {
                            return global::OfficeIMO.OfficeLength.Points((firstLine) / 20D);
                        }
                    }
                }
                return null;
            }
            set {
                var wordNativeValue = checked((int?)(value?.ToTwips()));
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                var indentation = props.Indentation ?? new Indentation();
                indentation.FirstLine = wordNativeValue.HasValue ? new StringValue(wordNativeValue.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)) : null;
                props.Indentation = indentation;
            }
        }
        /// <summary>
        /// Gets or sets the hanging indentation as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? IndentationHanging {
            get {
                if (_paragraphProperties != null && _paragraphProperties.Indentation != null) {
                    if (!string.IsNullOrEmpty(_paragraphProperties.Indentation.Hanging)) {
                        if (int.TryParse(_paragraphProperties.Indentation.Hanging, out var hanging)) {
                            return global::OfficeIMO.OfficeLength.Points((hanging) / 20D);
                        }
                    }
                }
                return null;
            }
            set {
                var wordNativeValue = checked((int?)(value?.ToTwips()));
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                var indentation = props.Indentation ?? new Indentation();
                indentation.Hanging = wordNativeValue.HasValue ? new StringValue(wordNativeValue.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)) : null;
                props.Indentation = indentation;
            }
        }
        /// <summary>
        /// Gets or sets the text flow direction for the paragraph.
        /// </summary>
        public WordTextDirection? TextDirection {
            get {
                if (_paragraphProperties != null && _paragraphProperties.TextDirection != null) {
                    return _paragraphProperties.TextDirection.Val?.Value.ToOfficeEnum();
                }
                return null;
            }
            set {
                var textDirection = new TextDirection { Val = value?.ToOpenXml() };
                if (_paragraphProperties == null) {
                    _paragraph.ParagraphProperties = new ParagraphProperties();
                }
                _paragraph.ParagraphProperties!.TextDirection = textDirection;
            }
        }

        /// <summary>
        /// Gets or sets directly authored right-to-left layout without resolving paragraph styles.
        /// </summary>
        public bool BiDi {
            get {
                return ReadParagraphOnOff(_paragraphProperties?.BiDi);
            }
            set {
                if (_paragraphProperties == null) {
                    _paragraph.ParagraphProperties = new ParagraphProperties();
                }

                var paragraphProps = _paragraph.ParagraphProperties!;

                if (value) {
                    paragraphProps.BiDi ??= new BiDi();
                } else {
                    paragraphProps.BiDi?.Remove();
                }
            }
        }

        private static bool ReadParagraphOnOff(OnOffType? value) =>
            value?.Val?.Value ?? value is not null;

        internal static bool HasPageBreakBefore(Paragraph paragraph) =>
            ReadParagraphOnOff(paragraph.ParagraphProperties?.PageBreakBefore);

        /// <summary>
        /// Gets or sets directly authored line spacing as a multiplier, exact height, or minimum height.
        /// Null removes both line attributes, retaining before/after and other spacing attributes.
        /// </summary>
        /// <remarks>
        /// This reads direct formatting, not effective style inheritance. Missing, incomplete or
        /// unsupported native declarations return null and remain unchanged until explicitly replaced.
        /// An omitted rule with a numeric line value uses automatic line multiples.
        /// </remarks>
        public WordLineSpacing? LineSpacing {
            get => WordLineSpacing.Read(_paragraph.ParagraphProperties?.SpacingBetweenLines);
            set {
                if (!value.HasValue) {
                    var existing = _paragraph.ParagraphProperties?.SpacingBetweenLines;
                    if (existing != null) { existing.Line = null; existing.LineRule = null; }
                    return;
                }
                // Validate and quantize before changing the document.
                string line = value.Value.ToNativeValue().ToString(System.Globalization.CultureInfo.InvariantCulture);
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                var spacing = props.SpacingBetweenLines ??= new SpacingBetweenLines();
                spacing.Line = line;
                spacing.LineRule = value.Value.Rule.ToOpenXml();
            }
        }

        // Native declaration access is intentionally internal. Codecs and native fixtures must
        // retain rule-only/value-only declarations rather than synthesize effective formatting.
        internal int? NativeLineSpacingValue {
            get => int.TryParse(_paragraph.ParagraphProperties?.SpacingBetweenLines?.Line?.Value,
                System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture,
                out int line) ? line : null;
            set {
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                var spacing = props.SpacingBetweenLines ??= new SpacingBetweenLines();
                spacing.Line = value?.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        internal WordLineSpacingRule? NativeLineSpacingRule {
            get {
                string? rule = _paragraph.ParagraphProperties?.SpacingBetweenLines?.LineRule?.InnerText;
                return rule == "auto" ? WordLineSpacingRule.Auto : rule == "exact" ? WordLineSpacingRule.Exact
                    : rule == "atLeast" ? WordLineSpacingRule.AtLeast : null;
            }
            set {
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                var spacing = props.SpacingBetweenLines ??= new SpacingBetweenLines();
                spacing.LineRule = value?.ToOpenXml();
            }
        }

        /// <summary>
        /// Gets or sets the spacing before the paragraph as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? LineSpacingBefore {
            get {
                if (_paragraphProperties != null && _paragraphProperties.SpacingBetweenLines != null) {
                    if (!string.IsNullOrEmpty(_paragraphProperties.SpacingBetweenLines.Before)) {
                        if (int.TryParse(_paragraphProperties.SpacingBetweenLines.Before, out var before)) {
                            return global::OfficeIMO.OfficeLength.Points((before) / 20D);
                        }
                    }
                }
                return null;
            }
            set {
                var wordNativeValue = checked((int?)(value?.ToTwips()));
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                var spacing = props.SpacingBetweenLines ?? new SpacingBetweenLines();
                spacing.Before = wordNativeValue?.ToString();
                props.SpacingBetweenLines = spacing;
            }
        }
        /// <summary>
        /// Gets or sets the spacing after the paragraph as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? LineSpacingAfter {
            get {
                if (_paragraphProperties != null && _paragraphProperties.SpacingBetweenLines != null) {
                    if (!string.IsNullOrEmpty(_paragraphProperties.SpacingBetweenLines.After)) {
                        if (int.TryParse(_paragraphProperties.SpacingBetweenLines.After, out var after)) {
                            return global::OfficeIMO.OfficeLength.Points((after) / 20D);
                        }
                    }
                }
                return null;
            }
            set {
                var wordNativeValue = checked((int?)(value?.ToTwips()));
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                var spacing = props.SpacingBetweenLines ?? new SpacingBetweenLines();
                spacing.After = wordNativeValue?.ToString();
                props.SpacingBetweenLines = spacing;
            }
        }

        /// <summary>
        /// Gets or sets the vertical text alignment - the alignment of the text in the paragraph with respect to the line height.
        /// </summary>
        public WordVerticalTextPosition? VerticalTextAlignment {
            get {
                RunProperties? runProperties = ScopedRunProperties;
                if (runProperties?.VerticalTextAlignment != null) {
                    return runProperties.VerticalTextAlignment.Val?.Value.ToOfficeEnum();
                }
                return null;
            }
            set {
                RunProperties? runProperties = ScopedRunProperties;
                if (value == null) {
                    if (runProperties?.VerticalTextAlignment == null) {
                        return;
                    }
                    runProperties.VerticalTextAlignment = null;
                } else {
                    if (IsHyperLink && _stdRun == null) {
                        WordHyperLink hyperlink = Hyperlink!;
                        runProperties = VerifyRunProperties(
                            hyperlink._hyperlink!,
                            hyperlink._run!,
                            hyperlink._runProperties);
                    } else {
                        runProperties = VerifyRunProperties();
                    }
                    runProperties.VerticalTextAlignment = new VerticalTextAlignment { Val = value.Value.ToOpenXml() };
                }
            }
        }
    }
}
