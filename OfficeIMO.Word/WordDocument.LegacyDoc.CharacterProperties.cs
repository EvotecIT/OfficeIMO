using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>Projects the character properties supported by Word styles.</summary>
        private static StyleRunProperties? CreateLegacyDocStyleRunProperties(LegacyDocCharacterFormat characterFormat, LegacyDocCharacterFormatProperties styleToggles = LegacyDocCharacterFormatProperties.None) {
            RunProperties? runProperties = CreateLegacyDocCharacterRunProperties(characterFormat, styleToggles);
            if (runProperties == null) return null;

            var properties = new StyleRunProperties();
            foreach (var property in runProperties.ChildElements) {
                if (property is not Highlight) properties.Append(property.CloneNode(true));
            }
            return properties.HasChildren ? properties : null;
        }

        /// <summary>Builds fresh imported character properties in WordprocessingML schema order.</summary>
        private static RunProperties? CreateLegacyDocCharacterRunProperties(LegacyDocCharacterFormat characterFormat, LegacyDocCharacterFormatProperties styleToggles = LegacyDocCharacterFormatProperties.None) {
            bool Resolve(bool value, LegacyDocCharacterFormatProperties property) =>
                ResolveLegacyDocToggle(value, property, characterFormat.StyleRelative, characterFormat.StyleInverted, styleToggles);

            var properties = new RunProperties();
            bool hasProperties = false;

            if (!string.IsNullOrEmpty(characterFormat.FontFamily)) {
                properties.Append(new RunFonts {
                    Ascii = characterFormat.FontFamily,
                    HighAnsi = characterFormat.FontFamily,
                    ComplexScript = characterFormat.FontFamily,
                    EastAsia = characterFormat.FontFamily
                });
                hasProperties = true;
            }

            hasProperties |= AppendLegacyDocCharacterOnOffProperty<Bold>(properties, Resolve(characterFormat.Bold, LegacyDocCharacterFormatProperties.Bold), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Bold));
            hasProperties |= AppendLegacyDocCharacterOnOffProperty<BoldComplexScript>(properties, Resolve(characterFormat.Bold, LegacyDocCharacterFormatProperties.Bold), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Bold));
            hasProperties |= AppendLegacyDocCharacterOnOffProperty<Italic>(properties, Resolve(characterFormat.Italic, LegacyDocCharacterFormatProperties.Italic), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Italic));
            hasProperties |= AppendLegacyDocCharacterOnOffProperty<ItalicComplexScript>(properties, Resolve(characterFormat.Italic, LegacyDocCharacterFormatProperties.Italic), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Italic));
            hasProperties |= AppendLegacyDocCharacterOnOffProperty<Caps>(properties, Resolve(characterFormat.Caps == LegacyDocCapsKind.Caps, LegacyDocCharacterFormatProperties.Caps), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Caps));
            hasProperties |= AppendLegacyDocCharacterOnOffProperty<SmallCaps>(properties, Resolve(characterFormat.Caps == LegacyDocCapsKind.SmallCaps, LegacyDocCharacterFormatProperties.SmallCaps), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.SmallCaps));
            hasProperties |= AppendLegacyDocCharacterOnOffProperty<Strike>(properties, Resolve(characterFormat.Strike, LegacyDocCharacterFormatProperties.Strike), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Strike));
            hasProperties |= AppendLegacyDocCharacterOnOffProperty<DoubleStrike>(properties, Resolve(characterFormat.DoubleStrike, LegacyDocCharacterFormatProperties.DoubleStrike), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.DoubleStrike));
            hasProperties |= AppendLegacyDocCharacterOnOffProperty<Outline>(properties, Resolve(characterFormat.Outline, LegacyDocCharacterFormatProperties.Outline), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Outline));
            hasProperties |= AppendLegacyDocCharacterOnOffProperty<Shadow>(properties, Resolve(characterFormat.Shadow, LegacyDocCharacterFormatProperties.Shadow), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Shadow));
            hasProperties |= AppendLegacyDocCharacterOnOffProperty<Emboss>(properties, Resolve(characterFormat.Emboss, LegacyDocCharacterFormatProperties.Emboss), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Emboss));
            hasProperties |= AppendLegacyDocCharacterOnOffProperty<Imprint>(properties, Resolve(characterFormat.Imprint, LegacyDocCharacterFormatProperties.Imprint), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Imprint));
            hasProperties |= AppendLegacyDocCharacterOnOffProperty<NoProof>(properties, Resolve(characterFormat.NoProof, LegacyDocCharacterFormatProperties.NoProof), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.NoProof));
            hasProperties |= AppendLegacyDocCharacterOnOffProperty<Vanish>(properties, Resolve(characterFormat.Hidden, LegacyDocCharacterFormatProperties.Hidden), characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Hidden));

            if (!string.IsNullOrEmpty(characterFormat.ColorHex)) {
                properties.Append(new Color { Val = characterFormat.ColorHex! });
                hasProperties = true;
            }

            if (characterFormat.CharacterSpacingTwips != null || characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.CharacterSpacing)) {
                properties.Append(new Spacing { Val = characterFormat.CharacterSpacingTwips ?? 0 });
                hasProperties = true;
            }

            if (characterFormat.CharacterScalePercentage.HasValue) {
                properties.Append(new CharacterScale { Val = characterFormat.CharacterScalePercentage.Value });
                hasProperties = true;
            }

            if (characterFormat.KerningMinimumFontSizeHalfPoints.HasValue) {
                properties.Append(new Kern { Val = (uint)characterFormat.KerningMinimumFontSizeHalfPoints.Value });
                hasProperties = true;
            }

            if (characterFormat.FontSizeHalfPoints != null) {
                string fontSize = characterFormat.FontSizeHalfPoints.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                properties.Append(new FontSize { Val = fontSize });
                properties.Append(new FontSizeComplexScript { Val = fontSize });
                hasProperties = true;
            }

            if (characterFormat.Highlight != null && TryMapHighlight(characterFormat.Highlight.Value, out HighlightColorValues highlight)) {
                properties.Append(new Highlight { Val = highlight });
                hasProperties = true;
            } else if (characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Highlight)) {
                properties.Append(new Highlight { Val = HighlightColorValues.None });
                hasProperties = true;
            }

            if (characterFormat.Underline != null && TryMapUnderline(characterFormat.Underline.Value, out UnderlineValues underline)) {
                properties.Append(new Underline { Val = underline });
                hasProperties = true;
            } else if (characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Underline)) {
                properties.Append(new Underline { Val = UnderlineValues.None });
                hasProperties = true;
            }

            if (characterFormat.VerticalPosition != null && TryMapVerticalPosition(characterFormat.VerticalPosition.Value, out VerticalPositionValues verticalPosition)) {
                properties.Append(new VerticalTextAlignment { Val = verticalPosition });
                hasProperties = true;
            } else if (characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.VerticalPosition)) {
                properties.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Baseline });
                hasProperties = true;
            }

            if (!string.IsNullOrEmpty(characterFormat.Language) || !string.IsNullOrEmpty(characterFormat.EastAsiaLanguage)) {
                properties.Append(CreateLegacyDocLanguages(characterFormat.Language, characterFormat.EastAsiaLanguage));
                hasProperties = true;
            }

            return hasProperties ? properties : null;
        }

        /// <summary>Appends one fresh character toggle, retaining explicit false overrides.</summary>
        private static bool AppendLegacyDocCharacterOnOffProperty<T>(RunProperties properties, bool enabled, bool specified) where T : OnOffType, new() {
            if (!enabled && !specified) {
                return false;
            }

            var property = new T();
            if (!enabled) {
                property.Val = false;
            }

            properties.Append(property);
            return true;
        }

    }
}
