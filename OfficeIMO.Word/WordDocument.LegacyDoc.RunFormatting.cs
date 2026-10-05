using OfficeIMO.Word.LegacyDoc.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private static void ApplyLegacyDocRunOnOffProperty<T>(WordParagraph run, bool enabled, bool specified) where T : OnOffType, new() {
            if (!enabled && !specified) {
                return;
            }

            RunProperties runProperties = run._runProperties ?? new RunProperties();
            run._runProperties = runProperties;
            runProperties.RemoveAllChildren<T>();

            var property = new T();
            if (!enabled) {
                property.Val = false;
            }

            runProperties.AddChild(property, true);
        }

        private static void ApplyLegacyDocRunFormatting(WordParagraph run, LegacyDocTextRun legacyRun) {
            LegacyDocCharacterFormatProperties styleToggles = legacyRun.StyleRelative == LegacyDocCharacterFormatProperties.None
                ? LegacyDocCharacterFormatProperties.None : GetLegacyDocStyleToggles(run);
            bool Resolve(bool value, LegacyDocCharacterFormatProperties property) =>
                ResolveLegacyDocToggle(value, property, legacyRun.StyleRelative, legacyRun.StyleInverted, styleToggles);

            ApplyLegacyDocRunOnOffProperty<Bold>(run, Resolve(legacyRun.Bold, LegacyDocCharacterFormatProperties.Bold), legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Bold));
            ApplyLegacyDocRunOnOffProperty<BoldComplexScript>(run, Resolve(legacyRun.Bold, LegacyDocCharacterFormatProperties.Bold), legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Bold));
            ApplyLegacyDocRunOnOffProperty<Italic>(run, Resolve(legacyRun.Italic, LegacyDocCharacterFormatProperties.Italic), legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Italic));
            ApplyLegacyDocRunOnOffProperty<ItalicComplexScript>(run, Resolve(legacyRun.Italic, LegacyDocCharacterFormatProperties.Italic), legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Italic));
            ApplyLegacyDocRunOnOffProperty<Strike>(run, Resolve(legacyRun.Strike, LegacyDocCharacterFormatProperties.Strike), legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Strike));
            ApplyLegacyDocRunOnOffProperty<DoubleStrike>(run, Resolve(legacyRun.DoubleStrike, LegacyDocCharacterFormatProperties.DoubleStrike), legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.DoubleStrike));
            ApplyLegacyDocRunOnOffProperty<Outline>(run, Resolve(legacyRun.Outline, LegacyDocCharacterFormatProperties.Outline), legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Outline));
            ApplyLegacyDocRunOnOffProperty<Shadow>(run, Resolve(legacyRun.Shadow, LegacyDocCharacterFormatProperties.Shadow), legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Shadow));
            ApplyLegacyDocRunOnOffProperty<Emboss>(run, Resolve(legacyRun.Emboss, LegacyDocCharacterFormatProperties.Emboss), legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Emboss));
            ApplyLegacyDocRunOnOffProperty<Imprint>(run, Resolve(legacyRun.Imprint, LegacyDocCharacterFormatProperties.Imprint), legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Imprint));
            ApplyLegacyDocRunOnOffProperty<Vanish>(run, Resolve(legacyRun.Hidden, LegacyDocCharacterFormatProperties.Hidden), legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Hidden));
            ApplyLegacyDocRunOnOffProperty<NoProof>(run, Resolve(legacyRun.NoProof, LegacyDocCharacterFormatProperties.NoProof), legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.NoProof));
            ApplyLegacyDocRunOnOffProperty<Caps>(run, Resolve(legacyRun.Caps == LegacyDocCapsKind.Caps, LegacyDocCharacterFormatProperties.Caps), legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Caps));
            ApplyLegacyDocRunOnOffProperty<SmallCaps>(run, Resolve(legacyRun.Caps == LegacyDocCapsKind.SmallCaps, LegacyDocCharacterFormatProperties.SmallCaps), legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.SmallCaps));

            if (legacyRun.VerticalPosition != null && TryMapVerticalPosition(legacyRun.VerticalPosition.Value, out VerticalPositionValues verticalPosition)) {
                run.VerticalTextAlignment = verticalPosition.ToOfficeEnum();
            } else if (legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.VerticalPosition)) {
                run.VerticalTextAlignment = WordVerticalTextPosition.Baseline;
            }

            if (legacyRun.Underline != null && TryMapUnderline(legacyRun.Underline.Value, out UnderlineValues underline)) {
                run.Underline = underline.ToOfficeEnum();
            } else if (legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Underline)) {
                run.Underline = WordUnderlineStyle.None;
            }

            if (legacyRun.Highlight != null && TryMapHighlight(legacyRun.Highlight.Value, out HighlightColorValues highlight)) {
                run.Highlight = highlight.ToOfficeEnum();
            } else if (legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Highlight)) {
                run.Highlight = WordHighlightColor.None;
            }

            if (legacyRun.FontSizeHalfPoints != null) {
                string fontSize = legacyRun.FontSizeHalfPoints.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                RunProperties runProperties = run._runProperties ?? new RunProperties();
                run._runProperties = runProperties;
                runProperties.FontSize = new FontSize {
                    Val = fontSize
                };
                runProperties.FontSizeComplexScript = new FontSizeComplexScript {
                    Val = fontSize
                };
            }

            if (!string.IsNullOrEmpty(legacyRun.ColorHex)) {
                run.ColorHex = legacyRun.ColorHex!;
            }

            if (!string.IsNullOrEmpty(legacyRun.FontFamily)) {
                run.SetFontFamily(legacyRun.FontFamily!);
            }

            if (legacyRun.KerningMinimumFontSizeHalfPoints.HasValue) {
                run.KerningMinimumFontSizePoints = legacyRun.KerningMinimumFontSizeHalfPoints.Value / 2D;
            }

            if (legacyRun.CharacterSpacingTwips != null || legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.CharacterSpacing)) {
                run.Spacing = legacyRun.CharacterSpacingTwips ?? 0;
            }

            if (!string.IsNullOrEmpty(legacyRun.Language) || !string.IsNullOrEmpty(legacyRun.EastAsiaLanguage)) {
                RunProperties runProperties = run._runProperties ?? new RunProperties();
                run._runProperties = runProperties;
                runProperties.Languages = CreateLegacyDocLanguages(legacyRun.Language, legacyRun.EastAsiaLanguage);
            }
        }

        private static bool ResolveLegacyDocToggle(bool value, LegacyDocCharacterFormatProperties property,
            LegacyDocCharacterFormatProperties relative, LegacyDocCharacterFormatProperties inverted,
            LegacyDocCharacterFormatProperties inherited) => (relative & property) == 0 ? value
                : ((inherited & property) != 0) ^ ((inverted & property) != 0);

        private static LegacyDocCharacterFormatProperties GetLegacyDocStyleToggles(WordParagraph run) {
            Styles? styles = run._document._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            if (styles == null) return LegacyDocCharacterFormatProperties.None;
            var seen = LegacyDocCharacterFormatProperties.None;
            var values = LegacyDocCharacterFormatProperties.None;
            ReadStyleChain(run._runProperties?.RunStyle?.Val?.Value);
            ReadStyleChain(run._paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "Normal");
            ReadProperties(styles.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle);
            return values;

            void ReadStyleChain(string? styleId) {
                var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                while (!string.IsNullOrEmpty(styleId) && visited.Add(styleId!)) {
                    Style? style = styles.Elements<Style>().FirstOrDefault(candidate => string.Equals(candidate.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase));
                    if (style == null) break;
                    ReadProperties(style.StyleRunProperties);
                    styleId = style.BasedOn?.Val?.Value;
                }
            }

            void ReadProperties(OpenXmlCompositeElement? properties) {
                if (properties == null) return;
                foreach (OnOffType toggle in properties.Elements<OnOffType>()) {
                    LegacyDocCharacterFormatProperties property;
                    switch (toggle.LocalName) {
                        case "b": property = LegacyDocCharacterFormatProperties.Bold; break;
                        case "i": property = LegacyDocCharacterFormatProperties.Italic; break;
                        case "strike": property = LegacyDocCharacterFormatProperties.Strike; break;
                        case "dstrike": property = LegacyDocCharacterFormatProperties.DoubleStrike; break;
                        case "outline": property = LegacyDocCharacterFormatProperties.Outline; break;
                        case "shadow": property = LegacyDocCharacterFormatProperties.Shadow; break;
                        case "emboss": property = LegacyDocCharacterFormatProperties.Emboss; break;
                        case "imprint": property = LegacyDocCharacterFormatProperties.Imprint; break;
                        case "vanish": property = LegacyDocCharacterFormatProperties.Hidden; break;
                        case "noProof": property = LegacyDocCharacterFormatProperties.NoProof; break;
                        case "caps": property = LegacyDocCharacterFormatProperties.Caps; break;
                        case "smallCaps": property = LegacyDocCharacterFormatProperties.SmallCaps; break;
                        default: continue;
                    }
                    if ((seen & property) != 0) continue;
                    seen |= property;
                    if (toggle.Val?.Value ?? true) values |= property;
                }
            }
        }

        private static LegacyDocCharacterFormatProperties GetLegacyDocParentStyleToggles(LegacyDocParagraphStyle source, LegacyDocStyleSheet sheet) {
            var chain = new Stack<LegacyDocCharacterFormat>();
            var visited = new HashSet<ushort> { source.Index };
            ushort? parent = source.BasedOnStyleIndex ?? (source.Index == 0 ? (ushort?)null : 0);
            while (parent.HasValue && visited.Add(parent.Value) && sheet.TryGetParagraphStyle(parent.Value, out LegacyDocParagraphStyle style)) {
                chain.Push(style.CharacterFormat);
                parent = style.BasedOnStyleIndex ?? (style.Index == 0 ? (ushort?)null : 0);
            }
            var values = LegacyDocCharacterFormatProperties.None;
            while (chain.Count > 0) {
                LegacyDocCharacterFormat format = chain.Pop();
                Apply(LegacyDocCharacterFormatProperties.Bold, format.Bold);
                Apply(LegacyDocCharacterFormatProperties.Italic, format.Italic);
                Apply(LegacyDocCharacterFormatProperties.Strike, format.Strike);
                Apply(LegacyDocCharacterFormatProperties.DoubleStrike, format.DoubleStrike);
                Apply(LegacyDocCharacterFormatProperties.Outline, format.Outline);
                Apply(LegacyDocCharacterFormatProperties.Shadow, format.Shadow);
                Apply(LegacyDocCharacterFormatProperties.Emboss, format.Emboss);
                Apply(LegacyDocCharacterFormatProperties.Imprint, format.Imprint);
                Apply(LegacyDocCharacterFormatProperties.Hidden, format.Hidden);
                Apply(LegacyDocCharacterFormatProperties.NoProof, format.NoProof);
                Apply(LegacyDocCharacterFormatProperties.Caps, format.Caps == LegacyDocCapsKind.Caps);
                Apply(LegacyDocCharacterFormatProperties.SmallCaps, format.Caps == LegacyDocCapsKind.SmallCaps);
                void Apply(LegacyDocCharacterFormatProperties property, bool value) {
                    if (!format.IsSpecified(property) && !value) return;
                    bool enabled = ResolveLegacyDocToggle(value, property, format.StyleRelative, format.StyleInverted, values);
                    values = enabled ? values | property : values & ~property;
                }
            }
            return values;
        }
    }
}
