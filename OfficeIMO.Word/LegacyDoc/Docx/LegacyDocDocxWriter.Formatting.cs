using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.LegacyDoc.Model;
using static OfficeIMO.Word.LegacyDoc.Docx.LegacyDocXml;

namespace OfficeIMO.Word.LegacyDoc.Docx {
    internal sealed partial class LegacyDocDocxWriter {
        private const int MaxStyleToggleDepth = 256;
        private Dictionary<string, int>? _styleIndex;
        private readonly Dictionary<(string? Run, string? Paragraph), LegacyDocCharacterFormatProperties> _styleToggleCache = new();

        /// <summary>
        /// Establishes the binary DOC paragraph baseline before importing source styles and direct formatting. Omitted
        /// DOC spacing means single lines and no paragraph spacing, rather than the DOCX authoring defaults.
        /// </summary>
        private void ApplyParagraphDefaults() {
            OfficeXmlTree styles = _styles;
            int defaults = Child(styles, styles.Root, "docDefaults");
            if (defaults == OfficeXmlTree.NoNode) {
                defaults = Element(styles, "docDefaults");
                styles.InsertBefore(styles.Root, defaults, styles.FirstChildOf(styles.Root));
            }

            int paragraphDefaults = Child(styles, defaults, "pPrDefault");
            if (paragraphDefaults == OfficeXmlTree.NoNode) paragraphDefaults = Append(styles, defaults, Element(styles, "pPrDefault"));
            int properties = Child(styles, paragraphDefaults, "pPr");
            if (properties == OfficeXmlTree.NoNode) properties = Append(styles, paragraphDefaults, Element(styles, "pPr"));
            int spacing = Element(styles, "spacing");
            Set(styles, spacing, "before", "0");
            Set(styles, spacing, "after", "0");
            Set(styles, spacing, "line", "240");
            Set(styles, spacing, "lineRule", "auto");
            Replace(styles, properties, spacing, ParagraphPropertiesOrder);
        }

        // ---- Runs -----------------------------------------------------------------------------------------------

        /// <summary>The run properties of <paramref name="run"/>, created as its first child when missing.</summary>
        private static int RunProperties(OfficeXmlTree tree, int run) {
            int properties = Child(tree, run, "rPr");
            if (properties != OfficeXmlTree.NoNode) return properties;
            properties = Element(tree, "rPr");
            tree.InsertBefore(run, properties, tree.FirstChildOf(run));
            return properties;
        }

        private void ApplyRunFormatting(RunTarget target, int run, LegacyDocTextRun legacyRun) {
            OfficeXmlTree tree = target.Tree;
            LegacyDocCharacterFormatProperties styleToggles = legacyRun.StyleRelative == LegacyDocCharacterFormatProperties.None
                ? LegacyDocCharacterFormatProperties.None
                : GetStyleToggles(tree, target.Paragraph, run);
            bool Resolve(bool value, LegacyDocCharacterFormatProperties property) =>
                ResolveToggle(value, property, legacyRun.StyleRelative, legacyRun.StyleInverted, styleToggles);

            void Toggle(string name, bool enabled, LegacyDocCharacterFormatProperties property) {
                if (!enabled && !legacyRun.IsSpecified(property)) return;
                Replace(tree, RunProperties(tree, run), OnOff(tree, name, enabled), RunPropertiesOrder);
            }

            bool bold = Resolve(legacyRun.Bold, LegacyDocCharacterFormatProperties.Bold);
            bool italic = Resolve(legacyRun.Italic, LegacyDocCharacterFormatProperties.Italic);
            Toggle("b", bold, LegacyDocCharacterFormatProperties.Bold);
            Toggle("bCs", bold, LegacyDocCharacterFormatProperties.Bold);
            Toggle("i", italic, LegacyDocCharacterFormatProperties.Italic);
            Toggle("iCs", italic, LegacyDocCharacterFormatProperties.Italic);
            Toggle("caps", Resolve(legacyRun.Caps == LegacyDocCapsKind.Caps, LegacyDocCharacterFormatProperties.Caps), LegacyDocCharacterFormatProperties.Caps);
            Toggle("smallCaps", Resolve(legacyRun.Caps == LegacyDocCapsKind.SmallCaps, LegacyDocCharacterFormatProperties.SmallCaps), LegacyDocCharacterFormatProperties.SmallCaps);
            Toggle("strike", Resolve(legacyRun.Strike, LegacyDocCharacterFormatProperties.Strike), LegacyDocCharacterFormatProperties.Strike);
            Toggle("dstrike", Resolve(legacyRun.DoubleStrike, LegacyDocCharacterFormatProperties.DoubleStrike), LegacyDocCharacterFormatProperties.DoubleStrike);
            Toggle("outline", Resolve(legacyRun.Outline, LegacyDocCharacterFormatProperties.Outline), LegacyDocCharacterFormatProperties.Outline);
            Toggle("shadow", Resolve(legacyRun.Shadow, LegacyDocCharacterFormatProperties.Shadow), LegacyDocCharacterFormatProperties.Shadow);
            Toggle("emboss", Resolve(legacyRun.Emboss, LegacyDocCharacterFormatProperties.Emboss), LegacyDocCharacterFormatProperties.Emboss);
            Toggle("imprint", Resolve(legacyRun.Imprint, LegacyDocCharacterFormatProperties.Imprint), LegacyDocCharacterFormatProperties.Imprint);
            Toggle("noProof", Resolve(legacyRun.NoProof, LegacyDocCharacterFormatProperties.NoProof), LegacyDocCharacterFormatProperties.NoProof);
            Toggle("vanish", Resolve(legacyRun.Hidden, LegacyDocCharacterFormatProperties.Hidden), LegacyDocCharacterFormatProperties.Hidden);

            string? verticalPosition = VerticalPositionToken(legacyRun.VerticalPosition);
            if (verticalPosition == null && legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.VerticalPosition)) verticalPosition = "baseline";
            if (verticalPosition != null) Replace(tree, RunProperties(tree, run), Element(tree, "vertAlign", verticalPosition), RunPropertiesOrder);

            string? underline = UnderlineToken(legacyRun.Underline);
            if (underline == null && legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Underline)) underline = "none";
            if (underline != null) Replace(tree, RunProperties(tree, run), Element(tree, "u", underline), RunPropertiesOrder);

            string? highlight = HighlightToken(legacyRun.Highlight);
            if (highlight == null && legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Highlight)) highlight = "none";
            if (highlight != null) Replace(tree, RunProperties(tree, run), Element(tree, "highlight", highlight), RunPropertiesOrder);

            if (legacyRun.FontSizeHalfPoints != null) {
                string fontSize = Text(legacyRun.FontSizeHalfPoints.Value);
                int properties = RunProperties(tree, run);
                Replace(tree, properties, Element(tree, "sz", fontSize), RunPropertiesOrder);
                Replace(tree, properties, Element(tree, "szCs", fontSize), RunPropertiesOrder);
            }

            if (!string.IsNullOrEmpty(legacyRun.ColorHex)) {
                Replace(tree, RunProperties(tree, run), Element(tree, "color", legacyRun.ColorHex!), RunPropertiesOrder);
            }

            if (!string.IsNullOrEmpty(legacyRun.FontFamily)) {
                Replace(tree, RunProperties(tree, run), CreateRunFonts(tree, legacyRun.FontFamily!), RunPropertiesOrder);
            }

            if (legacyRun.KerningMinimumFontSizeHalfPoints.HasValue) {
                Replace(tree, RunProperties(tree, run), Element(tree, "kern", Text(legacyRun.KerningMinimumFontSizeHalfPoints.Value)), RunPropertiesOrder);
            }

            if (legacyRun.CharacterScalePercentage.HasValue) {
                Replace(tree, RunProperties(tree, run), Element(tree, "w", Text(legacyRun.CharacterScalePercentage.Value)), RunPropertiesOrder);
            }

            if (legacyRun.CharacterSpacingTwips != null || legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.CharacterSpacing)) {
                Replace(tree, RunProperties(tree, run), Element(tree, "spacing", Text(legacyRun.CharacterSpacingTwips ?? 0)), RunPropertiesOrder);
            }

            if (!string.IsNullOrEmpty(legacyRun.Language) || !string.IsNullOrEmpty(legacyRun.EastAsiaLanguage)) {
                Replace(tree, RunProperties(tree, run), CreateLanguages(tree, legacyRun.Language, legacyRun.EastAsiaLanguage), RunPropertiesOrder);
            }
        }

        private static int CreateRunFonts(OfficeXmlTree tree, string fontFamily) {
            int fonts = Element(tree, "rFonts");
            Set(tree, fonts, "ascii", fontFamily);
            Set(tree, fonts, "hAnsi", fontFamily);
            Set(tree, fonts, "eastAsia", fontFamily);
            Set(tree, fonts, "cs", fontFamily);
            return fonts;
        }

        private static int CreateLanguages(OfficeXmlTree tree, string? language, string? eastAsiaLanguage) {
            int languages = Element(tree, "lang");
            if (!string.IsNullOrEmpty(language)) Set(tree, languages, "val", language);
            if (!string.IsNullOrEmpty(eastAsiaLanguage)) Set(tree, languages, "eastAsia", eastAsiaLanguage);
            return languages;
        }

        /// <summary>Builds fresh imported character properties in WordprocessingML schema order, or returns NoNode.</summary>
        private static int CreateCharacterRunProperties(OfficeXmlTree tree, LegacyDocCharacterFormat characterFormat, LegacyDocCharacterFormatProperties styleToggles, string elementName = "rPr") {
            bool Resolve(bool value, LegacyDocCharacterFormatProperties property) =>
                ResolveToggle(value, property, characterFormat.StyleRelative, characterFormat.StyleInverted, styleToggles);

            int properties = Element(tree, elementName);
            if (!string.IsNullOrEmpty(characterFormat.FontFamily)) Append(tree, properties, CreateRunFonts(tree, characterFormat.FontFamily!));

            void Toggle(string name, bool enabled, LegacyDocCharacterFormatProperties property) {
                if (!enabled && !characterFormat.IsSpecified(property)) return;
                Append(tree, properties, OnOff(tree, name, enabled));
            }

            bool bold = Resolve(characterFormat.Bold, LegacyDocCharacterFormatProperties.Bold);
            bool italic = Resolve(characterFormat.Italic, LegacyDocCharacterFormatProperties.Italic);
            Toggle("b", bold, LegacyDocCharacterFormatProperties.Bold);
            Toggle("bCs", bold, LegacyDocCharacterFormatProperties.Bold);
            Toggle("i", italic, LegacyDocCharacterFormatProperties.Italic);
            Toggle("iCs", italic, LegacyDocCharacterFormatProperties.Italic);
            Toggle("caps", Resolve(characterFormat.Caps == LegacyDocCapsKind.Caps, LegacyDocCharacterFormatProperties.Caps), LegacyDocCharacterFormatProperties.Caps);
            Toggle("smallCaps", Resolve(characterFormat.Caps == LegacyDocCapsKind.SmallCaps, LegacyDocCharacterFormatProperties.SmallCaps), LegacyDocCharacterFormatProperties.SmallCaps);
            Toggle("strike", Resolve(characterFormat.Strike, LegacyDocCharacterFormatProperties.Strike), LegacyDocCharacterFormatProperties.Strike);
            Toggle("dstrike", Resolve(characterFormat.DoubleStrike, LegacyDocCharacterFormatProperties.DoubleStrike), LegacyDocCharacterFormatProperties.DoubleStrike);
            Toggle("outline", Resolve(characterFormat.Outline, LegacyDocCharacterFormatProperties.Outline), LegacyDocCharacterFormatProperties.Outline);
            Toggle("shadow", Resolve(characterFormat.Shadow, LegacyDocCharacterFormatProperties.Shadow), LegacyDocCharacterFormatProperties.Shadow);
            Toggle("emboss", Resolve(characterFormat.Emboss, LegacyDocCharacterFormatProperties.Emboss), LegacyDocCharacterFormatProperties.Emboss);
            Toggle("imprint", Resolve(characterFormat.Imprint, LegacyDocCharacterFormatProperties.Imprint), LegacyDocCharacterFormatProperties.Imprint);
            Toggle("noProof", Resolve(characterFormat.NoProof, LegacyDocCharacterFormatProperties.NoProof), LegacyDocCharacterFormatProperties.NoProof);
            Toggle("vanish", Resolve(characterFormat.Hidden, LegacyDocCharacterFormatProperties.Hidden), LegacyDocCharacterFormatProperties.Hidden);

            if (!string.IsNullOrEmpty(characterFormat.ColorHex)) Append(tree, properties, Element(tree, "color", characterFormat.ColorHex!));
            if (characterFormat.CharacterSpacingTwips != null || characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.CharacterSpacing)) {
                Append(tree, properties, Element(tree, "spacing", Text(characterFormat.CharacterSpacingTwips ?? 0)));
            }

            if (characterFormat.CharacterScalePercentage.HasValue) Append(tree, properties, Element(tree, "w", Text(characterFormat.CharacterScalePercentage.Value)));
            if (characterFormat.KerningMinimumFontSizeHalfPoints.HasValue) Append(tree, properties, Element(tree, "kern", Text(characterFormat.KerningMinimumFontSizeHalfPoints.Value)));
            if (characterFormat.FontSizeHalfPoints != null) {
                string fontSize = Text(characterFormat.FontSizeHalfPoints.Value);
                Append(tree, properties, Element(tree, "sz", fontSize));
                Append(tree, properties, Element(tree, "szCs", fontSize));
            }

            string? highlight = HighlightToken(characterFormat.Highlight);
            if (highlight == null && characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Highlight)) highlight = "none";
            if (highlight != null) Append(tree, properties, Element(tree, "highlight", highlight));

            string? underline = UnderlineToken(characterFormat.Underline);
            if (underline == null && characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Underline)) underline = "none";
            if (underline != null) Append(tree, properties, Element(tree, "u", underline));

            string? verticalPosition = VerticalPositionToken(characterFormat.VerticalPosition);
            if (verticalPosition == null && characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.VerticalPosition)) verticalPosition = "baseline";
            if (verticalPosition != null) Append(tree, properties, Element(tree, "vertAlign", verticalPosition));

            if (!string.IsNullOrEmpty(characterFormat.Language) || !string.IsNullOrEmpty(characterFormat.EastAsiaLanguage)) {
                Append(tree, properties, CreateLanguages(tree, characterFormat.Language, characterFormat.EastAsiaLanguage));
            }

            return HasElements(tree, properties) ? properties : OfficeXmlTree.NoNode;
        }

        /// <summary>The character properties a style supports: everything except highlight.</summary>
        private static int CreateStyleRunProperties(OfficeXmlTree tree, LegacyDocCharacterFormat characterFormat, LegacyDocCharacterFormatProperties styleToggles = LegacyDocCharacterFormatProperties.None) {
            int properties = CreateCharacterRunProperties(tree, characterFormat, styleToggles);
            if (properties == OfficeXmlTree.NoNode) return OfficeXmlTree.NoNode;
            RemoveAll(tree, properties, "highlight");
            return HasElements(tree, properties) ? properties : OfficeXmlTree.NoNode;
        }

        // ---- Style toggles --------------------------------------------------------------------------------------

        private void ResetStyleToggleIndex() {
            _styleIndex = null;
            _styleToggleCache.Clear();
        }

        private static bool ResolveToggle(bool value, LegacyDocCharacterFormatProperties property,
            LegacyDocCharacterFormatProperties relative, LegacyDocCharacterFormatProperties inverted,
            LegacyDocCharacterFormatProperties inherited) => (relative & property) == 0 ? value
                : ((inherited & property) != 0) ^ ((inverted & property) != 0);

        private Dictionary<string, int> StyleIndex() {
            if (_styleIndex != null) return _styleIndex;
            var index = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (int style in Elements(_styles, _styles.Root, "style")) {
                string? id = Get(_styles, style, "styleId");
                if (!string.IsNullOrWhiteSpace(id) && !index.ContainsKey(id!)) index.Add(id!, style);
            }

            return _styleIndex = index;
        }

        /// <summary>The toggles the run's character style, its paragraph style and the document defaults turn on.</summary>
        private LegacyDocCharacterFormatProperties GetStyleToggles(OfficeXmlTree tree, int paragraph, int run) {
            Dictionary<string, int> index = StyleIndex();
            string? runStyleId = null;
            if (run != OfficeXmlTree.NoNode) {
                int runProperties = Child(tree, run, "rPr");
                int runStyle = runProperties == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : Child(tree, runProperties, "rStyle");
                if (runStyle != OfficeXmlTree.NoNode) runStyleId = Get(tree, runStyle, "val");
            }

            string? paragraphStyleId = ParagraphStyleId(tree, paragraph) ?? "Normal";
            var cacheKey = (runStyleId, paragraphStyleId);
            if (_styleToggleCache.TryGetValue(cacheKey, out LegacyDocCharacterFormatProperties cached)) return cached;
            var seen = LegacyDocCharacterFormatProperties.None;
            var values = LegacyDocCharacterFormatProperties.None;
            ReadStyleChain(runStyleId);
            ReadStyleChain(paragraphStyleId);
            int defaults = Child(_styles, _styles.Root, "docDefaults");
            int runDefaults = defaults == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : Child(_styles, defaults, "rPrDefault");
            if (runDefaults != OfficeXmlTree.NoNode) ReadProperties(Child(_styles, runDefaults, "rPr"));
            _styleToggleCache.Add(cacheKey, values);
            return values;

            void ReadStyleChain(string? styleId) {
                var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                while (!string.IsNullOrEmpty(styleId) && visited.Add(styleId!)) {
                    if (visited.Count > MaxStyleToggleDepth) throw new InvalidDataException("The legacy DOC style toggle inheritance chain exceeds the supported depth.");
                    if (!index.TryGetValue(styleId!, out int style)) break;
                    ReadProperties(Child(_styles, style, "rPr"));
                    int basedOn = Child(_styles, style, "basedOn");
                    styleId = basedOn == OfficeXmlTree.NoNode ? null : Get(_styles, basedOn, "val");
                }
            }

            void ReadProperties(int properties) {
                if (properties == OfficeXmlTree.NoNode) return;
                foreach (int toggle in Elements(_styles, properties)) {
                    if (_styles.NamespaceOf(toggle) != W) continue;
                    LegacyDocCharacterFormatProperties property;
                    switch (_styles.LocalNameOf(toggle)) {
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
                    if (IsOn(Get(_styles, toggle, "val"))) values |= property;
                }
            }
        }

        private static bool IsOn(string? value) =>
            value == null || !(value == "0" || string.Equals(value, "false", StringComparison.OrdinalIgnoreCase) || string.Equals(value, "off", StringComparison.OrdinalIgnoreCase));

        private static string? ParagraphStyleId(OfficeXmlTree tree, int paragraph) {
            int properties = Child(tree, paragraph, "pPr");
            if (properties == OfficeXmlTree.NoNode) return null;
            int style = Child(tree, properties, "pStyle");
            return style == OfficeXmlTree.NoNode ? null : Get(tree, style, "val");
        }

        private static LegacyDocCharacterFormatProperties GetParentStyleToggles(LegacyDocParagraphStyle source, LegacyDocStyleSheet sheet) {
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
                    bool enabled = ResolveToggle(value, property, format.StyleRelative, format.StyleInverted, values);
                    values = enabled ? values | property : values & ~property;
                }
            }

            return values;
        }

        // ---- Paragraphs -----------------------------------------------------------------------------------------

        /// <summary>The paragraph properties of <paramref name="paragraph"/>, created as its first child when missing.</summary>
        private static int ParagraphProperties(OfficeXmlTree tree, int paragraph) {
            int properties = Child(tree, paragraph, "pPr");
            if (properties != OfficeXmlTree.NoNode) return properties;
            properties = Element(tree, "pPr");
            tree.InsertBefore(paragraph, properties, tree.FirstChildOf(paragraph));
            return properties;
        }

        private void SetParagraphProperty(OfficeXmlTree tree, int paragraph, int property) =>
            Replace(tree, ParagraphProperties(tree, paragraph), property, ParagraphPropertiesOrder);

        private void ApplyParagraphFormatting(LegacyDocDocxPart part, int paragraph, LegacyDocParagraphFormat paragraphFormat) {
            OfficeXmlTree tree = part.Tree;
            if (paragraphFormat.StyleIndex != null) ApplyParagraphStyle(tree, paragraph, paragraphFormat.StyleIndex.Value);

            if (paragraphFormat.ParagraphMarkFormat != null) {
                int markProperties = CreateCharacterRunProperties(tree, paragraphFormat.ParagraphMarkFormat.Value, GetStyleToggles(tree, paragraph, OfficeXmlTree.NoNode));
                if (markProperties != OfficeXmlTree.NoNode) SetParagraphProperty(tree, paragraph, markProperties);
            }

            if (paragraphFormat.NumberingListIndex != null) {
                EnsureNumberingDefinition(paragraphFormat.NumberingListIndex.Value);
                int properties = ParagraphProperties(tree, paragraph);
                int numbering = Child(tree, properties, "numPr");
                if (numbering == OfficeXmlTree.NoNode) numbering = Append(tree, properties, Element(tree, "numPr"));
                ReplaceNumberingProperties(tree, numbering, paragraphFormat.NumberingListIndex.Value, paragraphFormat.NumberingLevel ?? 0);
            }

            string? alignment = ParagraphAlignmentToken(paragraphFormat.Alignment);
            if (alignment != null) SetParagraphProperty(tree, paragraph, Element(tree, "jc", alignment));

            string? verticalAlignment = VerticalCharacterAlignmentToken(paragraphFormat.VerticalCharacterAlignment);
            if (verticalAlignment != null) SetParagraphProperty(tree, paragraph, Element(tree, "textAlignment", verticalAlignment));

            if (paragraphFormat.OutlineLevel != null) SetParagraphProperty(tree, paragraph, Element(tree, "outlineLvl", Text(paragraphFormat.OutlineLevel.Value)));

            if (paragraphFormat.SpacingBeforeTwips != null || paragraphFormat.SpacingAfterTwips != null || paragraphFormat.LineSpacingTwips != null) {
                int spacing = GetOrAdd(tree, ParagraphProperties(tree, paragraph), "spacing", ParagraphPropertiesOrder);
                if (paragraphFormat.SpacingBeforeTwips != null) Set(tree, spacing, "before", Text(paragraphFormat.SpacingBeforeTwips.Value));
                if (paragraphFormat.SpacingAfterTwips != null) Set(tree, spacing, "after", Text(paragraphFormat.SpacingAfterTwips.Value));
                ApplyLineSpacing(tree, spacing, paragraphFormat);
            }

            if (paragraphFormat.LeftIndentTwips != null || paragraphFormat.RightIndentTwips != null || paragraphFormat.FirstLineIndentTwips != null) {
                int indentation = GetOrAdd(tree, ParagraphProperties(tree, paragraph), "ind", ParagraphPropertiesOrder);
                SetIndentation(tree, indentation, paragraphFormat);
            }

            SetParagraphToggle(tree, paragraph, "keepLines", paragraphFormat.KeepLinesTogether);
            SetParagraphToggle(tree, paragraph, "keepNext", paragraphFormat.KeepWithNext);
            SetParagraphToggle(tree, paragraph, "pageBreakBefore", paragraphFormat.PageBreakBefore);
            SetParagraphToggle(tree, paragraph, "widowControl", paragraphFormat.AvoidWidowAndOrphan);
            SetParagraphToggle(tree, paragraph, "suppressLineNumbers", paragraphFormat.SuppressLineNumbers);
            SetParagraphToggle(tree, paragraph, "suppressAutoHyphens", paragraphFormat.SuppressAutoHyphens);
            SetParagraphToggle(tree, paragraph, "contextualSpacing", paragraphFormat.ContextualSpacing);
            SetParagraphToggle(tree, paragraph, "mirrorIndents", paragraphFormat.MirrorIndents);
            SetParagraphToggle(tree, paragraph, "kinsoku", paragraphFormat.Kinsoku);
            SetParagraphToggle(tree, paragraph, "wordWrap", paragraphFormat.WordWrap);
            SetParagraphToggle(tree, paragraph, "overflowPunct", paragraphFormat.OverflowPunctuation);
            SetParagraphToggle(tree, paragraph, "topLinePunct", paragraphFormat.TopLinePunctuation);
            SetParagraphToggle(tree, paragraph, "autoSpaceDE", paragraphFormat.AutoSpaceDE);
            SetParagraphToggle(tree, paragraph, "autoSpaceDN", paragraphFormat.AutoSpaceDN);
            SetParagraphToggle(tree, paragraph, "bidi", paragraphFormat.Bidirectional);

            if (paragraphFormat.ParagraphShading != null && !string.IsNullOrEmpty(paragraphFormat.ParagraphShading.Value.FillColorHex)) {
                int shading = Element(tree, "shd");
                Set(tree, shading, "val", "clear");
                Set(tree, shading, "fill", paragraphFormat.ParagraphShading.Value.OpenXmlFill!);
                SetParagraphProperty(tree, paragraph, shading);
            }

            if (paragraphFormat.ParagraphBorders != null && paragraphFormat.ParagraphBorders.Value.HasAny) {
                ApplyParagraphBorders(tree, paragraph, paragraphFormat.ParagraphBorders.Value);
            }

            int tabs = CreateTabs(tree, paragraphFormat.TabStops);
            if (tabs != OfficeXmlTree.NoNode) {
                int existing = Child(tree, ParagraphProperties(tree, paragraph), "tabs");
                if (existing == OfficeXmlTree.NoNode) SetParagraphProperty(tree, paragraph, tabs);
                else foreach (int tab in Elements(tree, tabs)) {
                    tree.Remove(tab);
                    Append(tree, existing, tab);
                }
            }
        }

        private void SetParagraphToggle(OfficeXmlTree tree, int paragraph, string name, bool? value) {
            if (value.HasValue) SetParagraphProperty(tree, paragraph, OnOffValue(tree, name, value.Value));
        }

        private static void SetIndentation(OfficeXmlTree tree, int indentation, LegacyDocParagraphFormat paragraphFormat) {
            if (paragraphFormat.LeftIndentTwips != null) Set(tree, indentation, "left", Text(paragraphFormat.LeftIndentTwips.Value));
            if (paragraphFormat.RightIndentTwips != null) Set(tree, indentation, "right", Text(paragraphFormat.RightIndentTwips.Value));
            if (paragraphFormat.FirstLineIndentTwips != null) {
                if (paragraphFormat.FirstLineIndentTwips.Value < 0) {
                    tree.RemoveAttribute(indentation, W, "firstLine");
                    Set(tree, indentation, "hanging", Text(-paragraphFormat.FirstLineIndentTwips.Value));
                } else {
                    tree.RemoveAttribute(indentation, W, "hanging");
                    Set(tree, indentation, "firstLine", Text(paragraphFormat.FirstLineIndentTwips.Value));
                }
            }
        }

        /// <summary>Writes DOC LSPD's signed twips or 240ths of a line as the corresponding Open XML rule.</summary>
        private static void ApplyLineSpacing(OfficeXmlTree tree, int spacing, LegacyDocParagraphFormat format) {
            if (format.LineSpacingTwips is not int value) return;
            Set(tree, spacing, "line", Text(Math.Abs(value)));
            Set(tree, spacing, "lineRule", format.LineSpacingIsMultiple ? "auto" : value < 0 ? "exact" : "atLeast");
        }

        private void ApplyParagraphStyle(OfficeXmlTree tree, int paragraph, ushort styleIndex) {
            string? styleId = null;
            if (_styleSheet.TryGetParagraphStyle(styleIndex, out LegacyDocParagraphStyle legacyStyle)) {
                if (legacyStyle.BuiltInStyle != null) styleId = UseBuiltInStyle(legacyStyle.BuiltInStyle.Value.ToStringStyle());
                else if (!string.IsNullOrWhiteSpace(legacyStyle.StyleId)) styleId = legacyStyle.StyleId!;
            }

            if (styleId == null && TryMapBuiltInParagraphStyle(styleIndex, out WordParagraphStyles style)) styleId = UseBuiltInStyle(style.ToStringStyle());
            if (styleId == null) return;
            int properties = ParagraphProperties(tree, paragraph);
            int styleElement = Child(tree, properties, "pStyle");
            if (styleElement == OfficeXmlTree.NoNode) styleElement = Insert(tree, properties, Element(tree, "pStyle"), ParagraphPropertiesOrder);
            Set(tree, styleElement, "val", styleId);
        }

        /// <summary>
        /// The identifier to store for a built-in style: the package's own definition of it, found by identifier or by
        /// the built-in name, or the English identifier.
        /// </summary>
        private string UseBuiltInStyle(string builtInStyleId) {
            if (StyleIndex().ContainsKey(builtInStyleId)) return builtInStyleId;
            if (WordBuiltInStyleNames.TryGetName(builtInStyleId, out string builtInName)) {
                foreach (int style in Elements(_styles, _styles.Root, "style")) {
                    int name = Child(_styles, style, "name");
                    if (name != OfficeXmlTree.NoNode && string.Equals(Get(_styles, name, "val"), builtInName, StringComparison.OrdinalIgnoreCase)) {
                        return Get(_styles, style, "styleId") ?? builtInStyleId;
                    }
                }
            }

            return builtInStyleId;
        }

        private void ApplyParagraphBorders(OfficeXmlTree tree, int paragraph, LegacyDocParagraphBorders borders) {
            int properties = ParagraphProperties(tree, paragraph);
            int paragraphBorders = Child(tree, properties, "pBdr");
            bool created = false;
            if (paragraphBorders == OfficeXmlTree.NoNode) {
                paragraphBorders = Element(tree, "pBdr");
                created = true;
            }

            AddParagraphBorder(tree, paragraphBorders, "top", borders.Top, explicitSizeAndSpace: false);
            AddParagraphBorder(tree, paragraphBorders, "left", borders.Left, explicitSizeAndSpace: false);
            AddParagraphBorder(tree, paragraphBorders, "bottom", borders.Bottom, explicitSizeAndSpace: false);
            AddParagraphBorder(tree, paragraphBorders, "right", borders.Right, explicitSizeAndSpace: false);
            if (borders.Between.HasAny) AddParagraphBorder(tree, paragraphBorders, "between", borders.Between, explicitSizeAndSpace: true);
            if (created && HasElements(tree, paragraphBorders)) Insert(tree, properties, paragraphBorders, ParagraphPropertiesOrder);
        }

        /// <summary>
        /// Adds one paragraph or page border edge. Direct paragraph edges state only the size and space they have, as
        /// the paragraph border setters write them; style and page edges always state both.
        /// </summary>
        private static void AddParagraphBorder(OfficeXmlTree tree, int borders, string edge, LegacyDocParagraphBorder border, bool explicitSizeAndSpace) {
            string? style = ParagraphBorderStyleToken(border.Style);
            if (style == null) {
                if (!explicitSizeAndSpace) return;
                style = "nil";
            }

            int element = Element(tree, edge, style);
            if (!string.IsNullOrEmpty(border.ColorHex)) Set(tree, element, "color", border.ColorHex);
            if (explicitSizeAndSpace || border.SizeEighthPoints > 0) Set(tree, element, "sz", Text(border.SizeEighthPoints));
            if (explicitSizeAndSpace || border.SpacePoints > 0) Set(tree, element, "space", Text(border.SpacePoints));
            Replace(tree, borders, element, BordersOrder);
        }

        private static int CreateParagraphBorders(OfficeXmlTree tree, LegacyDocParagraphBorders borders, string elementName = "pBdr") {
            int paragraphBorders = Element(tree, elementName);
            if (borders.Top.HasAny) AddParagraphBorder(tree, paragraphBorders, "top", borders.Top, explicitSizeAndSpace: true);
            if (borders.Left.HasAny) AddParagraphBorder(tree, paragraphBorders, "left", borders.Left, explicitSizeAndSpace: true);
            if (borders.Bottom.HasAny) AddParagraphBorder(tree, paragraphBorders, "bottom", borders.Bottom, explicitSizeAndSpace: true);
            if (borders.Right.HasAny) AddParagraphBorder(tree, paragraphBorders, "right", borders.Right, explicitSizeAndSpace: true);
            if (borders.Between.HasAny) AddParagraphBorder(tree, paragraphBorders, "between", borders.Between, explicitSizeAndSpace: true);
            return paragraphBorders;
        }

        private static int CreateTabs(OfficeXmlTree tree, IReadOnlyList<LegacyDocTabStop> tabStops) {
            if (tabStops.Count == 0) return OfficeXmlTree.NoNode;
            int tabs = Element(tree, "tabs");
            foreach (LegacyDocTabStop tabStop in tabStops) {
                string? alignment = TabStopAlignmentToken(tabStop.Alignment);
                string? leader = TabStopLeaderToken(tabStop.Leader);
                if (alignment == null || leader == null) continue;
                int tab = Element(tree, "tab", alignment);
                Set(tree, tab, "leader", leader);
                Set(tree, tab, "pos", Text(tabStop.PositionTwips));
                Append(tree, tabs, tab);
            }

            return HasElements(tree, tabs) ? tabs : OfficeXmlTree.NoNode;
        }

        // ---- Value tokens ---------------------------------------------------------------------------------------

        private static bool TryMapBuiltInParagraphStyle(ushort styleIndex, out WordParagraphStyles style) {
            switch (styleIndex) {
                case 0: style = WordParagraphStyles.Normal; return true;
                case 1: style = WordParagraphStyles.Heading1; return true;
                case 2: style = WordParagraphStyles.Heading2; return true;
                case 3: style = WordParagraphStyles.Heading3; return true;
                case 4: style = WordParagraphStyles.Heading4; return true;
                case 5: style = WordParagraphStyles.Heading5; return true;
                case 6: style = WordParagraphStyles.Heading6; return true;
                case 7: style = WordParagraphStyles.Heading7; return true;
                case 8: style = WordParagraphStyles.Heading8; return true;
                case 9: style = WordParagraphStyles.Heading9; return true;
                default: style = default; return false;
            }
        }

        private static string? ParagraphAlignmentToken(LegacyDocParagraphAlignment? alignment) => alignment switch {
            LegacyDocParagraphAlignment.Left => "left",
            LegacyDocParagraphAlignment.Center => "center",
            LegacyDocParagraphAlignment.Right => "right",
            LegacyDocParagraphAlignment.Justify => "both",
            _ => null
        };

        private static string? VerticalCharacterAlignmentToken(byte? alignment) => alignment switch {
            0 => "top",
            1 => "center",
            2 => "baseline",
            3 => "bottom",
            4 => "auto",
            _ => null
        };

        private static string? VerticalPositionToken(LegacyDocVerticalPositionKind? position) => position switch {
            LegacyDocVerticalPositionKind.Superscript => "superscript",
            LegacyDocVerticalPositionKind.Subscript => "subscript",
            _ => null
        };

        private static string? HighlightToken(LegacyDocHighlightColorKind? highlight) => highlight switch {
            LegacyDocHighlightColorKind.Black => "black",
            LegacyDocHighlightColorKind.Blue => "blue",
            LegacyDocHighlightColorKind.Cyan => "cyan",
            LegacyDocHighlightColorKind.Green => "green",
            LegacyDocHighlightColorKind.Magenta => "magenta",
            LegacyDocHighlightColorKind.Red => "red",
            LegacyDocHighlightColorKind.Yellow => "yellow",
            LegacyDocHighlightColorKind.White => "white",
            LegacyDocHighlightColorKind.DarkBlue => "darkBlue",
            LegacyDocHighlightColorKind.DarkCyan => "darkCyan",
            LegacyDocHighlightColorKind.DarkGreen => "darkGreen",
            LegacyDocHighlightColorKind.DarkMagenta => "darkMagenta",
            LegacyDocHighlightColorKind.DarkRed => "darkRed",
            LegacyDocHighlightColorKind.DarkYellow => "darkYellow",
            LegacyDocHighlightColorKind.DarkGray => "darkGray",
            LegacyDocHighlightColorKind.LightGray => "lightGray",
            _ => null
        };

        private static string? UnderlineToken(LegacyDocUnderlineKind? underline) => underline switch {
            LegacyDocUnderlineKind.Single => "single",
            LegacyDocUnderlineKind.Words => "words",
            LegacyDocUnderlineKind.Double => "double",
            LegacyDocUnderlineKind.Dotted => "dotted",
            LegacyDocUnderlineKind.Thick => "thick",
            LegacyDocUnderlineKind.Dash => "dash",
            LegacyDocUnderlineKind.DotDash => "dotDash",
            LegacyDocUnderlineKind.DotDotDash => "dotDotDash",
            LegacyDocUnderlineKind.Wave => "wave",
            LegacyDocUnderlineKind.DottedHeavy => "dottedHeavy",
            LegacyDocUnderlineKind.DashedHeavy => "dashedHeavy",
            LegacyDocUnderlineKind.DashDotHeavy => "dashDotHeavy",
            LegacyDocUnderlineKind.DashDotDotHeavy => "dashDotDotHeavy",
            LegacyDocUnderlineKind.WavyHeavy => "wavyHeavy",
            LegacyDocUnderlineKind.DashLong => "dashLong",
            LegacyDocUnderlineKind.WavyDouble => "wavyDouble",
            LegacyDocUnderlineKind.DashLongHeavy => "dashLongHeavy",
            _ => null
        };

        private static string? TabStopAlignmentToken(LegacyDocTabStopAlignment alignment) => alignment switch {
            LegacyDocTabStopAlignment.Left => "left",
            LegacyDocTabStopAlignment.Center => "center",
            LegacyDocTabStopAlignment.Right => "right",
            LegacyDocTabStopAlignment.Decimal => "decimal",
            LegacyDocTabStopAlignment.Bar => "bar",
            LegacyDocTabStopAlignment.Number => "num",
            LegacyDocTabStopAlignment.Clear => "clear",
            _ => null
        };

        private static string? TabStopLeaderToken(LegacyDocTabStopLeader leader) => leader switch {
            LegacyDocTabStopLeader.None => "none",
            LegacyDocTabStopLeader.Dot => "dot",
            LegacyDocTabStopLeader.Hyphen => "hyphen",
            LegacyDocTabStopLeader.Underscore => "underscore",
            LegacyDocTabStopLeader.Heavy => "heavy",
            LegacyDocTabStopLeader.MiddleDot => "middleDot",
            _ => null
        };

        private static string? ParagraphBorderStyleToken(LegacyDocParagraphBorderStyle style) => style switch {
            LegacyDocParagraphBorderStyle.Single => "single",
            LegacyDocParagraphBorderStyle.Double => "double",
            LegacyDocParagraphBorderStyle.Dotted => "dotted",
            LegacyDocParagraphBorderStyle.Dashed => "dashed",
            _ => null
        };
    }
}
