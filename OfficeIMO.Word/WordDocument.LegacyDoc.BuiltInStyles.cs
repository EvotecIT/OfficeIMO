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
            MergeLegacyDocBuiltInStyleRunFormatting(style, legacyStyle.CharacterFormat);
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
                    spacing.Line = paragraphFormat.LineSpacingTwips.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    spacing.LineRule = LineSpacingRuleValues.AtLeast;
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

            if (paragraphFormat.KeepLinesTogether == true) {
                ReplaceStyleProperty(properties, new KeepLines());
            }

            if (paragraphFormat.KeepWithNext == true) {
                ReplaceStyleProperty(properties, new KeepNext());
            }

            if (paragraphFormat.PageBreakBefore == true) {
                ReplaceStyleProperty(properties, new PageBreakBefore());
            }

            if (paragraphFormat.AvoidWidowAndOrphan == true) {
                ReplaceStyleProperty(properties, new WidowControl());
            }

            if (paragraphFormat.SuppressLineNumbers == true) {
                ReplaceStyleProperty(properties, new SuppressLineNumbers());
            }

            if (paragraphFormat.SuppressAutoHyphens == true) {
                ReplaceStyleProperty(properties, new SuppressAutoHyphens());
            }

            if (paragraphFormat.ContextualSpacing == true) {
                ReplaceStyleProperty(properties, new ContextualSpacing());
            }

            if (paragraphFormat.MirrorIndents == true) {
                ReplaceStyleProperty(properties, new MirrorIndents());
            }

            if (paragraphFormat.Kinsoku == true) {
                ReplaceStyleProperty(properties, new Kinsoku());
            }

            if (paragraphFormat.WordWrap == true) {
                ReplaceStyleProperty(properties, new WordWrap());
            }

            if (paragraphFormat.OverflowPunctuation == true) {
                ReplaceStyleProperty(properties, new OverflowPunctuation());
            }

            if (paragraphFormat.TopLinePunctuation == true) {
                ReplaceStyleProperty(properties, new TopLinePunctuation());
            }

            if (paragraphFormat.AutoSpaceDE == true) {
                ReplaceStyleProperty(properties, new AutoSpaceDE());
            }

            if (paragraphFormat.AutoSpaceDN == true) {
                ReplaceStyleProperty(properties, new AutoSpaceDN());
            }

            if (paragraphFormat.Bidirectional == true) {
                ReplaceStyleProperty(properties, new BiDi());
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

        private static void MergeLegacyDocBuiltInStyleRunFormatting(Style style, LegacyDocCharacterFormat characterFormat) {
            if (!characterFormat.HasFormatting) {
                return;
            }

            StyleRunProperties properties = style.StyleRunProperties ?? style.AppendChild(new StyleRunProperties());

            if (!string.IsNullOrEmpty(characterFormat.FontFamily)) {
                ReplaceStyleProperty(properties, new RunFonts {
                    Ascii = characterFormat.FontFamily,
                    HighAnsi = characterFormat.FontFamily,
                    ComplexScript = characterFormat.FontFamily,
                    EastAsia = characterFormat.FontFamily
                });
            }

            ReplaceStyleOnOffProperty<Bold>(properties, characterFormat.Bold, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Bold));
            ReplaceStyleOnOffProperty<BoldComplexScript>(properties, characterFormat.Bold, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Bold));
            ReplaceStyleOnOffProperty<Italic>(properties, characterFormat.Italic, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Italic));
            ReplaceStyleOnOffProperty<ItalicComplexScript>(properties, characterFormat.Italic, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Italic));
            ReplaceStyleOnOffProperty<Strike>(properties, characterFormat.Strike, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Strike));
            ReplaceStyleOnOffProperty<DoubleStrike>(properties, characterFormat.DoubleStrike, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.DoubleStrike));
            ReplaceStyleOnOffProperty<Outline>(properties, characterFormat.Outline, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Outline));
            ReplaceStyleOnOffProperty<Shadow>(properties, characterFormat.Shadow, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Shadow));
            ReplaceStyleOnOffProperty<Emboss>(properties, characterFormat.Emboss, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Emboss));
            ReplaceStyleOnOffProperty<Imprint>(properties, characterFormat.Imprint, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Imprint));
            ReplaceStyleOnOffProperty<Vanish>(properties, characterFormat.Hidden, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Hidden));
            ReplaceStyleOnOffProperty<NoProof>(properties, characterFormat.NoProof, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.NoProof));
            ReplaceStyleOnOffProperty<Caps>(properties, characterFormat.Caps == LegacyDocCapsKind.Caps, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Caps));
            ReplaceStyleOnOffProperty<SmallCaps>(properties, characterFormat.Caps == LegacyDocCapsKind.SmallCaps, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.SmallCaps));

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
            }

            if (characterFormat.Underline != null && TryMapUnderline(characterFormat.Underline.Value, out UnderlineValues underline)) {
                ReplaceStyleProperty(properties, new Underline { Val = underline });
            }

            if (characterFormat.VerticalPosition != null && TryMapVerticalPosition(characterFormat.VerticalPosition.Value, out VerticalPositionValues verticalPosition)) {
                ReplaceStyleProperty(properties, new VerticalTextAlignment { Val = verticalPosition });
            }
        }

        private static void ReplaceStyleProperty<T>(OpenXmlCompositeElement parent, T replacement) where T : OpenXmlElement {
            RemoveStyleProperties<T>(parent);
            parent.Append(replacement);
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
