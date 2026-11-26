using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using static OfficeIMO.Word.OpenXml.WordXmlEdits;

namespace OfficeIMO.Word {
    // Provides accessors for paragraph formatting. Each reads and edits the paragraph properties through their element
    // handle, so it works the same whether OfficeIMO's own package model or the Open XML SDK holds the document.
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
        public string? StyleId => ChildValue(ParagraphPropertiesNode, "pStyle");

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
            get => ParagraphTokenValue<WordParagraphAlignment>("jc", WordParagraphAlignmentTokens.TryParse);
            set => SetParagraphToken("jc", value == null ? null : WordParagraphAlignmentTokens.ToToken(value.Value));
        }

        /// <summary>
        /// Text Alignment aka Vertical Character Alignment on Line. This element specifies the vertical alignment of all text on each line displayed within a paragraph. If the line height (before any added spacing) is larger than one or more characters on the line, all characters are aligned to each other as specified by this element.
        /// If this element is omitted on a given paragraph, its value is determined by the setting previously set at any level of the style hierarchy (i.e.that previous setting remains unchanged). If this setting is never specified in the style hierarchy, then the vertical alignment of all characters on the line shall be automatically determined by the consumer.
        /// Setting this property to <see langword="null"/> removes direct alignment and restores style inheritance.
        /// </summary>
        public WordVerticalCharacterAlignment? VerticalCharacterAlignmentOnLine {
            get => ParagraphTokenValue<WordVerticalCharacterAlignment>("textAlignment", WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue);
            set => SetParagraphToken("textAlignment", value?.ToOpenXmlToken());
        }

        /// <summary>
        /// Gets or sets the indentation before the paragraph as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? IndentationBefore {
            get => GetParagraphTwips("ind", "left");
            set => SetParagraphTwips("ind", "left", IndentationText(value));
        }

        /// <summary>
        /// Gets or sets the indentation after the paragraph as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? IndentationAfter {
            get => GetParagraphTwips("ind", "right");
            set => SetParagraphTwips("ind", "right", IndentationText(value));
        }

        /// <summary>Gets or sets a directly authored page-break-before value; this does not resolve inherited paragraph styles.</summary>
        public bool PageBreakBefore {
            get => ReadOnOff(ParagraphPropertiesNode.Child(W, "pageBreakBefore"));
            set => SetPageBreakBefore(value);
        }

        /// <summary>
        /// Gets or sets directly authored keep-with-next formatting without resolving paragraph styles.
        /// </summary>
        public bool KeepWithNext {
            get => ReadOnOff(ParagraphPropertiesNode.Child(W, "keepNext"));
            set => SetParagraphFlag("keepNext", value);
        }

        /// <summary>
        /// Gets or sets directly authored keep-lines-together formatting without resolving paragraph styles.
        /// </summary>
        public bool KeepLinesTogether {
            get => ReadOnOff(ParagraphPropertiesNode.Child(W, "keepLines"));
            set => SetParagraphFlag("keepLines", value);
        }

        /// <summary>
        /// Gets or sets directly authored widow and orphan control without resolving paragraph styles.
        /// </summary>
        public bool AvoidWidowAndOrphan {
            get => ReadOnOff(ParagraphPropertiesNode.Child(W, "widowControl"));
            set => SetWidowControl(value);
        }

        /// <summary>
        /// Gets or sets the first line indentation as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? IndentationFirstLine {
            get => GetParagraphTwips("ind", "firstLine");
            set => SetParagraphTwips("ind", "firstLine", IndentationText(value));
        }

        /// <summary>
        /// Gets or sets the hanging indentation as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? IndentationHanging {
            get => GetParagraphTwips("ind", "hanging");
            set => SetParagraphTwips("ind", "hanging", IndentationText(value));
        }

        /// <summary>
        /// Gets or sets the text flow direction for the paragraph.
        /// </summary>
        public WordTextDirection? TextDirection {
            get => ParagraphTokenValue<WordTextDirection>("textDirection", WordTextDirectionTokens.TryParse);
            set {
                // Setting no direction still writes the element, without a value.
                OfficeXmlNode direction = ReplaceChild(EditParagraphProperties(), "textDirection", WPPrView.Order);
                if (value != null) direction.SetAttribute(W, "val", WordTextDirectionTokens.ToToken(value.Value));
            }
        }

        /// <summary>
        /// Gets or sets directly authored right-to-left layout without resolving paragraph styles.
        /// </summary>
        public bool BiDi {
            get => ReadOnOff(ParagraphPropertiesNode.Child(W, "bidi"));
            set => SetBiDi(value);
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
                GetLineAttributes(out string? line, out string? rule);
                return WordLineSpacing.Read(line, rule);
            }
            set {
                // Validate and quantize before changing the document.
                string? line = value?.ToNativeValue().ToString(System.Globalization.CultureInfo.InvariantCulture);
                string? rule = value?.Rule.ToOpenXmlToken();
                SetLineAttributes(true, line, true, rule);
            }
        }

        // Native declaration access is intentionally internal. Codecs and native fixtures must
        // retain rule-only/value-only declarations rather than synthesize effective formatting.
        internal int? NativeLineSpacingValue {
            get {
                GetLineAttributes(out string? line, out _);
                return int.TryParse(line, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int value) ? value : null;
            }
            set {
                string? line = value?.ToString(System.Globalization.CultureInfo.InvariantCulture);
                SetLineAttributes(true, line, false, null);
            }
        }

        internal WordLineSpacingRule? NativeLineSpacingRule {
            get {
                GetLineAttributes(out _, out string? rule);
                return rule == "auto" ? WordLineSpacingRule.Auto : rule == "exact" ? WordLineSpacingRule.Exact
                    : rule == "atLeast" ? WordLineSpacingRule.AtLeast : null;
            }
            set {
                string? rule = value?.ToOpenXmlToken();
                SetLineAttributes(false, null, true, rule);
            }
        }

        /// <summary>
        /// Gets or sets the spacing before the paragraph as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? LineSpacingBefore {
            get => GetParagraphTwips("spacing", "before");
            set => SetParagraphTwips("spacing", "before", SpacingText(value));
        }

        /// <summary>
        /// Gets or sets the spacing after the paragraph as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? LineSpacingAfter {
            get => GetParagraphTwips("spacing", "after");
            set => SetParagraphTwips("spacing", "after", SpacingText(value));
        }

        /// <summary>
        /// Gets or sets the vertical text alignment - the alignment of the text in the paragraph with respect to the line height.
        /// </summary>
        public WordVerticalTextPosition? VerticalTextAlignment {
            get {
                string? token = ChildValue(RunPropertiesNode, "vertAlign");
                return token == null ? null : Token(WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordVerticalTextPosition parsed), parsed, token);
            }
            set {
                // Clearing leaves a run without the property as it is.
                if (value == null) {
                    if (!RunPropertiesNode.Child(W, "vertAlign").IsNone) RemoveChild(EditRunProperties(), "vertAlign");
                    return;
                }
                ReplaceChild(EditRunProperties(), "vertAlign", WRPrView.Order, value.Value.ToOpenXmlToken());
            }
        }
    }
}
