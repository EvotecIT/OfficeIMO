namespace OfficeIMO.Word {
    // Provides accessors for paragraph formatting. Each reads and edits OfficeIMO's own package model while the document
    // works without the Open XML SDK; once the SDK holds the document, the same member goes through the bridge.
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
        public string? StyleId => _ownedTree != null && TryGetOwnedStyleId(out string? owned) ? owned : SdkStyleId;

        /// <summary>
        /// The English identifier of the built-in style this paragraph uses, such as <c>Quote</c> for a French
        /// template's <c>Citation</c>, or <see langword="null"/> for a custom style or no style.
        /// </summary>
        internal string? BuiltInStyleId {
            get {
                string? styleId = StyleId;
                return _document != null ? _document.GetBuiltInStyleId(styleId) : WordBuiltInStyleNames.IsBuiltInId(styleId) ? styleId : null;
            }
        }

        /// <summary>
        /// Alignment aka Paragraph Alignment. This element specifies the paragraph alignment which shall be applied to text in this paragraph.
        /// If this element is omitted on a given paragraph, its value is determined by the setting previously set at any level of the style hierarchy (i.e.that previous setting remains unchanged). If this setting is never specified in the style hierarchy, then no alignment is applied to the paragraph.
        /// Setting this property to <see langword="null"/> removes direct alignment and restores style inheritance.
        /// </summary>
        public WordParagraphAlignment? ParagraphAlignment {
            get => _ownedTree != null && TryGetOwnedAlignment(out WordParagraphAlignment? owned) ? owned : SdkParagraphAlignment;
            set {
                if (_ownedTree != null && TrySetOwnedAlignment(value)) return;
                SdkParagraphAlignment = value;
            }
        }

        /// <summary>
        /// Text Alignment aka Vertical Character Alignment on Line. This element specifies the vertical alignment of all text on each line displayed within a paragraph. If the line height (before any added spacing) is larger than one or more characters on the line, all characters are aligned to each other as specified by this element.
        /// If this element is omitted on a given paragraph, its value is determined by the setting previously set at any level of the style hierarchy (i.e.that previous setting remains unchanged). If this setting is never specified in the style hierarchy, then the vertical alignment of all characters on the line shall be automatically determined by the consumer.
        /// Setting this property to <see langword="null"/> removes direct alignment and restores style inheritance.
        /// </summary>
        public WordVerticalCharacterAlignment? VerticalCharacterAlignmentOnLine {
            get => _ownedTree != null && TryGetOwnedVerticalCharacterAlignment(out WordVerticalCharacterAlignment? owned) ? owned : SdkVerticalCharacterAlignment;
            set {
                if (_ownedTree != null && TrySetOwnedVerticalCharacterAlignment(value)) return;
                SdkVerticalCharacterAlignment = value;
            }
        }

        /// <summary>
        /// Gets or sets the indentation before the paragraph as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? IndentationBefore {
            get => _ownedTree != null && TryGetOwnedParagraphTwips("ind", "left", out global::OfficeIMO.OfficeLength? owned) ? owned : GetSdkParagraphTwips("ind", "left");
            set {
                if (_ownedTree != null && TrySetOwnedParagraphTwips("ind", "left", IndentationText(value))) return;
                SetSdkParagraphTwips("ind", "left", value);
            }
        }

        /// <summary>
        /// Gets or sets the indentation after the paragraph as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? IndentationAfter {
            get => _ownedTree != null && TryGetOwnedParagraphTwips("ind", "right", out global::OfficeIMO.OfficeLength? owned) ? owned : GetSdkParagraphTwips("ind", "right");
            set {
                if (_ownedTree != null && TrySetOwnedParagraphTwips("ind", "right", IndentationText(value))) return;
                SetSdkParagraphTwips("ind", "right", value);
            }
        }

        /// <summary>Gets or sets a directly authored page-break-before value; this does not resolve inherited paragraph styles.</summary>
        public bool PageBreakBefore {
            get => _ownedTree != null && TryGetOwnedParagraphOnOff("pageBreakBefore", out bool owned) ? owned : GetSdkParagraphOnOff("pageBreakBefore");
            set {
                if (_ownedTree != null && TrySetOwnedPageBreakBefore(value)) return;
                SetSdkParagraphOnOff("pageBreakBefore", value);
            }
        }

        /// <summary>
        /// Gets or sets directly authored keep-with-next formatting without resolving paragraph styles.
        /// </summary>
        public bool KeepWithNext {
            get => _ownedTree != null && TryGetOwnedParagraphOnOff("keepNext", out bool owned) ? owned : GetSdkParagraphOnOff("keepNext");
            set {
                if (_ownedTree != null && TrySetOwnedParagraphFlag("keepNext", value)) return;
                SetSdkParagraphOnOff("keepNext", value);
            }
        }

        /// <summary>
        /// Gets or sets directly authored keep-lines-together formatting without resolving paragraph styles.
        /// </summary>
        public bool KeepLinesTogether {
            get => _ownedTree != null && TryGetOwnedParagraphOnOff("keepLines", out bool owned) ? owned : GetSdkParagraphOnOff("keepLines");
            set {
                if (_ownedTree != null && TrySetOwnedParagraphFlag("keepLines", value)) return;
                SetSdkParagraphOnOff("keepLines", value);
            }
        }

        /// <summary>
        /// Gets or sets directly authored widow and orphan control without resolving paragraph styles.
        /// </summary>
        public bool AvoidWidowAndOrphan {
            get => _ownedTree != null && TryGetOwnedParagraphOnOff("widowControl", out bool owned) ? owned : GetSdkParagraphOnOff("widowControl");
            set {
                if (_ownedTree != null && TrySetOwnedWidowControl(value)) return;
                SetSdkParagraphOnOff("widowControl", value);
            }
        }

        /// <summary>
        /// Gets or sets the first line indentation as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? IndentationFirstLine {
            get => _ownedTree != null && TryGetOwnedParagraphTwips("ind", "firstLine", out global::OfficeIMO.OfficeLength? owned) ? owned : GetSdkParagraphTwips("ind", "firstLine");
            set {
                if (_ownedTree != null && TrySetOwnedParagraphTwips("ind", "firstLine", IndentationText(value))) return;
                SetSdkParagraphTwips("ind", "firstLine", value);
            }
        }

        /// <summary>
        /// Gets or sets the hanging indentation as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? IndentationHanging {
            get => _ownedTree != null && TryGetOwnedParagraphTwips("ind", "hanging", out global::OfficeIMO.OfficeLength? owned) ? owned : GetSdkParagraphTwips("ind", "hanging");
            set {
                if (_ownedTree != null && TrySetOwnedParagraphTwips("ind", "hanging", IndentationText(value))) return;
                SetSdkParagraphTwips("ind", "hanging", value);
            }
        }

        /// <summary>
        /// Gets or sets the text flow direction for the paragraph.
        /// </summary>
        public WordTextDirection? TextDirection {
            get => _ownedTree != null && TryGetOwnedTextDirection(out WordTextDirection? owned) ? owned : SdkTextDirection;
            set {
                if (_ownedTree != null && TrySetOwnedTextDirection(value)) return;
                SdkTextDirection = value;
            }
        }

        /// <summary>
        /// Gets or sets directly authored right-to-left layout without resolving paragraph styles.
        /// </summary>
        public bool BiDi {
            get => _ownedTree != null && TryGetOwnedParagraphOnOff("bidi", out bool owned) ? owned : GetSdkParagraphOnOff("bidi");
            set {
                if (_ownedTree != null && TrySetOwnedBiDi(value)) return;
                SetSdkParagraphOnOff("bidi", value);
            }
        }

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
            get {
                if (_ownedTree != null && TryGetOwnedLineSpacing(out WordLineSpacing? owned)) return owned;
                GetSdkLineAttributes(out string? line, out string? rule);
                return WordLineSpacing.Read(line, rule);
            }
            set {
                // Validate and quantize before changing the document.
                string? line = value?.ToNativeValue().ToString(System.Globalization.CultureInfo.InvariantCulture);
                string? rule = value?.Rule.ToOpenXmlToken();
                if (_ownedTree != null && TrySetOwnedLineAttributes(true, line, true, rule)) return;
                SetSdkLineAttributes(true, line, true, rule);
            }
        }

        // Native declaration access is intentionally internal. Codecs and native fixtures must
        // retain rule-only/value-only declarations rather than synthesize effective formatting.
        internal int? NativeLineSpacingValue {
            get {
                string? line;
                if (_ownedTree == null || !TryGetOwnedLineAttributes(out line, out _)) GetSdkLineAttributes(out line, out _);
                return int.TryParse(line, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int value) ? value : null;
            }
            set {
                string? line = value?.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (_ownedTree != null && TrySetOwnedLineAttributes(true, line, false, null)) return;
                SetSdkLineAttributes(true, line, false, null);
            }
        }

        internal WordLineSpacingRule? NativeLineSpacingRule {
            get {
                string? rule;
                if (_ownedTree == null || !TryGetOwnedLineAttributes(out _, out rule)) GetSdkLineAttributes(out _, out rule);
                return rule == "auto" ? WordLineSpacingRule.Auto : rule == "exact" ? WordLineSpacingRule.Exact
                    : rule == "atLeast" ? WordLineSpacingRule.AtLeast : null;
            }
            set {
                string? rule = value?.ToOpenXmlToken();
                if (_ownedTree != null && TrySetOwnedLineAttributes(false, null, true, rule)) return;
                SetSdkLineAttributes(false, null, true, rule);
            }
        }

        /// <summary>
        /// Gets or sets the spacing before the paragraph as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? LineSpacingBefore {
            get => _ownedTree != null && TryGetOwnedParagraphTwips("spacing", "before", out global::OfficeIMO.OfficeLength? owned) ? owned : GetSdkParagraphTwips("spacing", "before");
            set {
                if (_ownedTree != null && TrySetOwnedParagraphTwips("spacing", "before", SpacingText(value))) return;
                SetSdkParagraphTwips("spacing", "before", value);
            }
        }

        /// <summary>
        /// Gets or sets the spacing after the paragraph as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? LineSpacingAfter {
            get => _ownedTree != null && TryGetOwnedParagraphTwips("spacing", "after", out global::OfficeIMO.OfficeLength? owned) ? owned : GetSdkParagraphTwips("spacing", "after");
            set {
                if (_ownedTree != null && TrySetOwnedParagraphTwips("spacing", "after", SpacingText(value))) return;
                SetSdkParagraphTwips("spacing", "after", value);
            }
        }

        /// <summary>
        /// Gets or sets the vertical text alignment - the alignment of the text in the paragraph with respect to the line height.
        /// </summary>
        public WordVerticalTextPosition? VerticalTextAlignment {
            get => _ownedTree != null && TryGetOwnedVerticalTextPosition(out WordVerticalTextPosition? owned) ? owned : SdkVerticalTextAlignment;
            set {
                if (_ownedTree != null && TrySetOwnedVerticalTextPosition(value)) return;
                SdkVerticalTextAlignment = value;
            }
        }
    }
}
