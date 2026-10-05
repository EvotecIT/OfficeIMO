using OfficeIMO.Word.LegacyDoc.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private static Style GetOrCreateLegacyDocBuiltInStyle(Styles styles, string styleId, string styleName) {
            Style? existing = styles
                .OfType<Style>()
                .FirstOrDefault(style => string.Equals(style.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase));
            if (existing != null) {
                if (existing.GetFirstChild<StyleName>() == null) {
                    existing.PrependChild(new StyleName { Val = styleName });
                }

                return existing;
            }

            var style = new Style { Type = StyleValues.Paragraph, StyleId = styleId };
            style.Append(new StyleName { Val = styleName });
            styles.Append(style);
            return style;
        }

        private static void MergeLegacyDocBuiltInStyleFormatting(WordDocument document, Style style, LegacyDocParagraphStyle legacyStyle, LegacyDocStyleSheet styleSheet) {
            MergeLegacyDocBuiltInStyleBasedOn(style, legacyStyle, styleSheet);
            MergeLegacyDocBuiltInStyleParagraphFormatting(document, style, legacyStyle.ParagraphFormat);
            MergeLegacyDocBuiltInStyleRunFormatting(style, legacyStyle.CharacterFormat, GetLegacyDocParentStyleToggles(legacyStyle, styleSheet));
        }

        private static void MergeLegacyDocBuiltInStyleBasedOn(Style style, LegacyDocParagraphStyle legacyStyle, LegacyDocStyleSheet styleSheet) {
            if (legacyStyle.BasedOnStyleIndex == null || legacyStyle.BasedOnStyleIndex.Value == legacyStyle.Index) {
                return;
            }

            string basedOnStyleId = ResolveLegacyDocBasedOnStyleId(legacyStyle, styleSheet);
            if (string.Equals(basedOnStyleId, style.StyleId?.Value, StringComparison.OrdinalIgnoreCase)) {
                return;
            }

            ReplaceStyleProperty(style, new BasedOn { Val = basedOnStyleId });
        }

        private static void MergeLegacyDocBuiltInStyleParagraphFormatting(WordDocument document, Style style, LegacyDocParagraphFormat paragraphFormat) {
            // Imported styles inherit absent controls from their own base style, not from our authoring templates.
            if (style.StyleParagraphProperties is StyleParagraphProperties templateProperties) {
                RemoveStyleProperties<KeepLines>(templateProperties);
                RemoveStyleProperties<KeepNext>(templateProperties);
                RemoveStyleProperties<PageBreakBefore>(templateProperties);
                RemoveStyleProperties<WidowControl>(templateProperties);
                RemoveStyleProperties<SuppressLineNumbers>(templateProperties);
                RemoveStyleProperties<SuppressAutoHyphens>(templateProperties);
                RemoveStyleProperties<ContextualSpacing>(templateProperties);
                RemoveStyleProperties<MirrorIndents>(templateProperties);
                RemoveStyleProperties<Kinsoku>(templateProperties);
                RemoveStyleProperties<WordWrap>(templateProperties);
                RemoveStyleProperties<OverflowPunctuation>(templateProperties);
                RemoveStyleProperties<TopLinePunctuation>(templateProperties);
                RemoveStyleProperties<AutoSpaceDE>(templateProperties);
                RemoveStyleProperties<AutoSpaceDN>(templateProperties);
                RemoveStyleProperties<BiDi>(templateProperties);
            }
            if (!paragraphFormat.HasFormatting) {
                return;
            }

            StyleParagraphProperties properties = style.StyleParagraphProperties ?? style.AppendChild(new StyleParagraphProperties());

            if (paragraphFormat.Alignment != null && TryMapParagraphAlignment(paragraphFormat.Alignment.Value, out JustificationValues alignment)) {
                ReplaceStyleProperty(properties, new Justification { Val = alignment });
            }

            if (paragraphFormat.NumberingListIndex != null) {
                ReplaceStyleProperty(properties, CreateLegacyDocNumberingProperties(document, paragraphFormat.NumberingListIndex.Value, paragraphFormat.NumberingLevel ?? 0));
            }

            if (paragraphFormat.SpacingBeforeTwips != null || paragraphFormat.SpacingAfterTwips != null || paragraphFormat.LineSpacingTwips != null) {
                SpacingBetweenLines spacing = properties.GetFirstChild<SpacingBetweenLines>() ?? properties.AppendChild(new SpacingBetweenLines());
                if (paragraphFormat.SpacingBeforeTwips != null) {
                    spacing.Before = paragraphFormat.SpacingBeforeTwips.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }

                if (paragraphFormat.SpacingAfterTwips != null) {
                    spacing.After = paragraphFormat.SpacingAfterTwips.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }

                if (paragraphFormat.LineSpacingTwips != null) {
                    ApplyLegacyDocLineSpacing(spacing, paragraphFormat);
                }
            }

            if (paragraphFormat.LeftIndentTwips != null || paragraphFormat.RightIndentTwips != null || paragraphFormat.FirstLineIndentTwips != null) {
                Indentation indentation = properties.GetFirstChild<Indentation>() ?? properties.AppendChild(new Indentation());
                if (paragraphFormat.LeftIndentTwips != null) {
                    indentation.Left = paragraphFormat.LeftIndentTwips.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }

                if (paragraphFormat.RightIndentTwips != null) {
                    indentation.Right = paragraphFormat.RightIndentTwips.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }

                if (paragraphFormat.FirstLineIndentTwips != null) {
                    if (paragraphFormat.FirstLineIndentTwips.Value < 0) {
                        indentation.Hanging = (-paragraphFormat.FirstLineIndentTwips.Value).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    } else {
                        indentation.FirstLine = paragraphFormat.FirstLineIndentTwips.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    }
                }
            }

            Tabs? tabs = CreateLegacyDocTabs(paragraphFormat.TabStops);
            if (tabs != null) {
                RemoveStyleProperties<Tabs>(properties);
                properties.Append(tabs);
            }

            if (paragraphFormat.KeepLinesTogether.HasValue) {
                ReplaceStyleProperty(properties, new KeepLines { Val = paragraphFormat.KeepLinesTogether.Value });
            }

            if (paragraphFormat.KeepWithNext.HasValue) {
                ReplaceStyleProperty(properties, new KeepNext { Val = paragraphFormat.KeepWithNext.Value });
            }

            if (paragraphFormat.PageBreakBefore.HasValue) {
                ReplaceStyleProperty(properties, new PageBreakBefore { Val = paragraphFormat.PageBreakBefore.Value });
            }

            if (paragraphFormat.AvoidWidowAndOrphan.HasValue) {
                ReplaceStyleProperty(properties, new WidowControl { Val = paragraphFormat.AvoidWidowAndOrphan.Value });
            }

            if (paragraphFormat.SuppressLineNumbers.HasValue) {
                ReplaceStyleProperty(properties, new SuppressLineNumbers { Val = paragraphFormat.SuppressLineNumbers.Value });
            }

            if (paragraphFormat.SuppressAutoHyphens.HasValue) {
                ReplaceStyleProperty(properties, new SuppressAutoHyphens { Val = paragraphFormat.SuppressAutoHyphens.Value });
            }

            if (paragraphFormat.ContextualSpacing.HasValue) {
                ReplaceStyleProperty(properties, new ContextualSpacing { Val = paragraphFormat.ContextualSpacing.Value });
            }

            if (paragraphFormat.MirrorIndents.HasValue) {
                ReplaceStyleProperty(properties, new MirrorIndents { Val = paragraphFormat.MirrorIndents.Value });
            }

            if (paragraphFormat.Kinsoku.HasValue) {
                ReplaceStyleProperty(properties, new Kinsoku { Val = paragraphFormat.Kinsoku.Value });
            }

            if (paragraphFormat.WordWrap.HasValue) {
                ReplaceStyleProperty(properties, new WordWrap { Val = paragraphFormat.WordWrap.Value });
            }

            if (paragraphFormat.OverflowPunctuation.HasValue) {
                ReplaceStyleProperty(properties, new OverflowPunctuation { Val = paragraphFormat.OverflowPunctuation.Value });
            }

            if (paragraphFormat.TopLinePunctuation.HasValue) {
                ReplaceStyleProperty(properties, new TopLinePunctuation { Val = paragraphFormat.TopLinePunctuation.Value });
            }

            if (paragraphFormat.AutoSpaceDE.HasValue) {
                ReplaceStyleProperty(properties, new AutoSpaceDE { Val = paragraphFormat.AutoSpaceDE.Value });
            }

            if (paragraphFormat.AutoSpaceDN.HasValue) {
                ReplaceStyleProperty(properties, new AutoSpaceDN { Val = paragraphFormat.AutoSpaceDN.Value });
            }

            if (paragraphFormat.Bidirectional.HasValue) {
                ReplaceStyleProperty(properties, new BiDi { Val = paragraphFormat.Bidirectional.Value });
            }

            if (paragraphFormat.VerticalCharacterAlignment != null && TryMapVerticalCharacterAlignment(paragraphFormat.VerticalCharacterAlignment.Value, out VerticalTextAlignmentValues verticalCharacterAlignment)) {
                ReplaceStyleProperty(properties, new TextAlignment { Val = verticalCharacterAlignment });
            }

            if (paragraphFormat.OutlineLevel != null) {
                ReplaceStyleProperty(properties, new OutlineLevel { Val = paragraphFormat.OutlineLevel.Value });
            }

            if (paragraphFormat.ParagraphShading != null && !string.IsNullOrEmpty(paragraphFormat.ParagraphShading.Value.FillColorHex)) {
                ReplaceStyleProperty(properties, new Shading {
                    Val = ShadingPatternValues.Clear,
                    Color = "auto",
                    Fill = paragraphFormat.ParagraphShading.Value.FillColorHex!
                });
            }

            if (paragraphFormat.ParagraphBorders != null && paragraphFormat.ParagraphBorders.Value.HasAny) {
                ReplaceStyleProperty(properties, CreateLegacyDocStyleParagraphBorders(paragraphFormat.ParagraphBorders.Value));
            }
        }

        private static void MergeLegacyDocBuiltInStyleRunFormatting(Style style, LegacyDocCharacterFormat characterFormat, LegacyDocCharacterFormatProperties styleToggles) {
            bool Resolve(bool value, LegacyDocCharacterFormatProperties property) =>
                ResolveLegacyDocToggle(value, property, characterFormat.StyleRelative, characterFormat.StyleInverted, styleToggles);

            if (style.StyleRunProperties is StyleRunProperties templateProperties) {
                // Missing source formatting inherits through basedOn. Authoring
                // template sizes, fonts, colors and effects must not override it.
                templateProperties.RemoveAllChildren();
            }
            if (!characterFormat.HasFormatting) {
                return;
            }

            StyleRunProperties properties = style.StyleRunProperties ?? style.AppendChild(new StyleRunProperties());

            if (characterFormat.KerningMinimumFontSizeHalfPoints.HasValue) {
                ReplaceStyleProperty(properties, new Kern { Val = (uint)characterFormat.KerningMinimumFontSizeHalfPoints.Value });
            }

            if (!string.IsNullOrEmpty(characterFormat.FontFamily)) {
                ReplaceStyleProperty(properties, new RunFonts {
                    Ascii = characterFormat.FontFamily,
                    HighAnsi = characterFormat.FontFamily,
                    ComplexScript = characterFormat.FontFamily,
                    EastAsia = characterFormat.FontFamily
                });
            }

            if (!string.IsNullOrEmpty(characterFormat.Language) || !string.IsNullOrEmpty(characterFormat.EastAsiaLanguage)) {
                ReplaceStyleProperty(properties, CreateLegacyDocLanguages(characterFormat.Language, characterFormat.EastAsiaLanguage));
            }

            ReplaceStyleOnOffProperty<Bold>(properties, Resolve(characterFormat.Bold, LegacyDocCharacterFormatProperties.Bold), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Bold));
            ReplaceStyleOnOffProperty<BoldComplexScript>(properties, Resolve(characterFormat.Bold, LegacyDocCharacterFormatProperties.Bold), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Bold));
            ReplaceStyleOnOffProperty<Italic>(properties, Resolve(characterFormat.Italic, LegacyDocCharacterFormatProperties.Italic), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Italic));
            ReplaceStyleOnOffProperty<ItalicComplexScript>(properties, Resolve(characterFormat.Italic, LegacyDocCharacterFormatProperties.Italic), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Italic));
            ReplaceStyleOnOffProperty<Strike>(properties, Resolve(characterFormat.Strike, LegacyDocCharacterFormatProperties.Strike), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Strike));
            ReplaceStyleOnOffProperty<DoubleStrike>(properties, Resolve(characterFormat.DoubleStrike, LegacyDocCharacterFormatProperties.DoubleStrike), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.DoubleStrike));
            ReplaceStyleOnOffProperty<Outline>(properties, Resolve(characterFormat.Outline, LegacyDocCharacterFormatProperties.Outline), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Outline));
            ReplaceStyleOnOffProperty<Shadow>(properties, Resolve(characterFormat.Shadow, LegacyDocCharacterFormatProperties.Shadow), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Shadow));
            ReplaceStyleOnOffProperty<Emboss>(properties, Resolve(characterFormat.Emboss, LegacyDocCharacterFormatProperties.Emboss), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Emboss));
            ReplaceStyleOnOffProperty<Imprint>(properties, Resolve(characterFormat.Imprint, LegacyDocCharacterFormatProperties.Imprint), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Imprint));
            ReplaceStyleOnOffProperty<Vanish>(properties, Resolve(characterFormat.Hidden, LegacyDocCharacterFormatProperties.Hidden), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Hidden));
            ReplaceStyleOnOffProperty<NoProof>(properties, Resolve(characterFormat.NoProof, LegacyDocCharacterFormatProperties.NoProof), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.NoProof));
            ReplaceStyleOnOffProperty<Caps>(properties, Resolve(characterFormat.Caps == LegacyDocCapsKind.Caps, LegacyDocCharacterFormatProperties.Caps), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Caps));
            ReplaceStyleOnOffProperty<SmallCaps>(properties, Resolve(characterFormat.Caps == LegacyDocCapsKind.SmallCaps, LegacyDocCharacterFormatProperties.SmallCaps), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.SmallCaps));

            if (!string.IsNullOrEmpty(characterFormat.ColorHex)) {
                ReplaceStyleProperty(properties, new Color { Val = characterFormat.ColorHex! });
            }

            if (characterFormat.FontSizeHalfPoints != null) {
                string fontSize = characterFormat.FontSizeHalfPoints.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                ReplaceStyleProperty(properties, new FontSize { Val = fontSize });
                ReplaceStyleProperty(properties, new FontSizeComplexScript { Val = fontSize });
            }

            if (characterFormat.Highlight != null && TryMapHighlight(characterFormat.Highlight.Value, out HighlightColorValues highlight)) {
                ReplaceStyleProperty(properties, new Highlight { Val = highlight });
            } else if (characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Highlight)) {
                ReplaceStyleProperty(properties, new Highlight { Val = HighlightColorValues.None });
            }

            if (characterFormat.Underline != null && TryMapUnderline(characterFormat.Underline.Value, out UnderlineValues underline)) {
                ReplaceStyleProperty(properties, new Underline { Val = underline });
            } else if (characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Underline)) {
                ReplaceStyleProperty(properties, new Underline { Val = UnderlineValues.None });
            }

            if (characterFormat.VerticalPosition != null && TryMapVerticalPosition(characterFormat.VerticalPosition.Value, out VerticalPositionValues verticalPosition)) {
                ReplaceStyleProperty(properties, new VerticalTextAlignment { Val = verticalPosition });
            } else if (characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.VerticalPosition)) {
                ReplaceStyleProperty(properties, new VerticalTextAlignment { Val = VerticalPositionValues.Baseline });
            }
        }

        private static void ReplaceStyleProperty<T>(OpenXmlCompositeElement parent, T replacement) where T : OpenXmlElement {
            RemoveStyleProperties<T>(parent);
            if (!parent.AddChild(replacement, false)) parent.Append(replacement);
        }

        private static void ReplaceStyleOnOffProperty<T>(OpenXmlCompositeElement parent, bool enabled, bool specified) where T : OnOffType, new() {
            if (!enabled && !specified) {
                return;
            }

            var replacement = new T();
            if (!enabled) {
                replacement.Val = false;
            }

            ReplaceStyleProperty(parent, replacement);
        }

        private static void RemoveStyleProperties<T>(OpenXmlCompositeElement parent) where T : OpenXmlElement {
            foreach (T child in parent.Elements<T>().ToArray()) {
                child.Remove();
            }
        }
    }
}
