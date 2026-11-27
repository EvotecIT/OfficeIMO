using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.LegacyDoc.Model;
using static OfficeIMO.Word.LegacyDoc.Docx.LegacyDocXml;

namespace OfficeIMO.Word.LegacyDoc.Docx {
    internal sealed partial class LegacyDocDocxWriter {
        private const string NumberingPartName = "/word/numbering.xml";
        private const string NumberingContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
        private OfficeXmlTree? _numbering;

        // ---- Paragraph styles -----------------------------------------------------------------------------------

        private void AddParagraphStyleDefinitions() {
            OfficeXmlTree styles = _styles;
            var existingStyleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (int style in Elements(styles, styles.Root, "style")) {
                string? styleId = Get(styles, style, "styleId");
                if (!string.IsNullOrWhiteSpace(styleId)) existingStyleIds.Add(styleId!);
            }

            foreach (LegacyDocParagraphStyle legacyStyle in _styleSheet.ParagraphStyles) {
                if (legacyStyle.BuiltInStyle != null) {
                    string builtInStyleId = legacyStyle.BuiltInStyle.Value.ToStringStyle();
                    int builtInStyle = GetOrCreateBuiltInStyle(builtInStyleId, legacyStyle.Name);
                    MergeBuiltInStyleBasedOn(builtInStyle, legacyStyle);
                    MergeBuiltInStyleParagraphFormatting(builtInStyle, legacyStyle.ParagraphFormat);
                    MergeBuiltInStyleRunFormatting(builtInStyle, legacyStyle.CharacterFormat, GetParentStyleToggles(legacyStyle, _styleSheet));
                    string? builtInNextStyleId = ResolveNextStyleId(legacyStyle);
                    if (!string.IsNullOrWhiteSpace(builtInNextStyleId)) Replace(styles, builtInStyle, Element(styles, "next", builtInNextStyleId!), StyleOrder);
                    continue;
                }

                if (string.IsNullOrWhiteSpace(legacyStyle.StyleId) || !existingStyleIds.Add(legacyStyle.StyleId!)) continue;

                int customStyle = Element(styles, "style");
                Set(styles, customStyle, "type", "paragraph");
                Set(styles, customStyle, "customStyle", "true");
                Set(styles, customStyle, "styleId", legacyStyle.StyleId);
                Append(styles, customStyle, Element(styles, "name", legacyStyle.Name));
                Append(styles, customStyle, Element(styles, "basedOn", ResolveBasedOnStyleId(legacyStyle)));
                string? nextStyleId = ResolveNextStyleId(legacyStyle);
                if (!string.IsNullOrWhiteSpace(nextStyleId)) Append(styles, customStyle, Element(styles, "next", nextStyleId!));
                int paragraphProperties = CreateStyleParagraphProperties(styles, legacyStyle.ParagraphFormat);
                if (paragraphProperties != OfficeXmlTree.NoNode) Append(styles, customStyle, paragraphProperties);
                int runProperties = CreateStyleRunProperties(styles, legacyStyle.CharacterFormat, GetParentStyleToggles(legacyStyle, _styleSheet));
                if (runProperties != OfficeXmlTree.NoNode) Append(styles, customStyle, runProperties);
                Append(styles, styles.Root, customStyle);
            }
        }

        private string ResolveBasedOnStyleId(LegacyDocParagraphStyle legacyStyle) {
            if (legacyStyle.BasedOnStyleIndex != null
                && legacyStyle.BasedOnStyleIndex.Value != legacyStyle.Index
                && _styleSheet.TryGetParagraphStyle(legacyStyle.BasedOnStyleIndex.Value, out LegacyDocParagraphStyle basedOnStyle)) {
                if (basedOnStyle.BuiltInStyle != null) return basedOnStyle.BuiltInStyle.Value.ToStringStyle();
                if (!string.IsNullOrWhiteSpace(basedOnStyle.StyleId)) return basedOnStyle.StyleId!;
            }

            if (legacyStyle.BasedOnStyleIndex != null && TryMapBuiltInParagraphStyle(legacyStyle.BasedOnStyleIndex.Value, out WordParagraphStyles builtInStyle)) {
                return builtInStyle.ToStringStyle();
            }

            return WordParagraphStyles.Normal.ToStringStyle();
        }

        private string? ResolveNextStyleId(LegacyDocParagraphStyle legacyStyle) {
            if (legacyStyle.NextStyleIndex == null || legacyStyle.NextStyleIndex.Value == legacyStyle.Index) return null;
            if (_styleSheet.TryGetParagraphStyle(legacyStyle.NextStyleIndex.Value, out LegacyDocParagraphStyle nextStyle)) {
                if (nextStyle.BuiltInStyle != null) return nextStyle.BuiltInStyle.Value.ToStringStyle();
                if (!string.IsNullOrWhiteSpace(nextStyle.StyleId)) return nextStyle.StyleId!;
            }

            return TryMapBuiltInParagraphStyle(legacyStyle.NextStyleIndex.Value, out WordParagraphStyles builtInStyle) ? builtInStyle.ToStringStyle() : null;
        }

        private int GetOrCreateBuiltInStyle(string styleId, string styleName) {
            OfficeXmlTree styles = _styles;
            foreach (int existing in Elements(styles, styles.Root, "style")) {
                if (!string.Equals(Get(styles, existing, "styleId"), styleId, StringComparison.OrdinalIgnoreCase)) continue;
                if (Child(styles, existing, "name") == OfficeXmlTree.NoNode) {
                    styles.InsertBefore(existing, Element(styles, "name", styleName), styles.FirstChildOf(existing));
                }

                return existing;
            }

            int style = Element(styles, "style");
            Set(styles, style, "type", "paragraph");
            Set(styles, style, "styleId", styleId);
            Append(styles, style, Element(styles, "name", styleName));
            Append(styles, styles.Root, style);
            return style;
        }

        private void MergeBuiltInStyleBasedOn(int style, LegacyDocParagraphStyle legacyStyle) {
            if (legacyStyle.BasedOnStyleIndex == null || legacyStyle.BasedOnStyleIndex.Value == legacyStyle.Index) return;
            string basedOnStyleId = ResolveBasedOnStyleId(legacyStyle);
            if (string.Equals(basedOnStyleId, Get(_styles, style, "styleId"), StringComparison.OrdinalIgnoreCase)) return;
            Replace(_styles, style, Element(_styles, "basedOn", basedOnStyleId), StyleOrder);
        }

        private static readonly string[] TemplateParagraphControls = {
            "spacing", "keepLines", "keepNext", "pageBreakBefore", "widowControl", "suppressLineNumbers", "suppressAutoHyphens",
            "contextualSpacing", "mirrorIndents", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi"
        };

        private void MergeBuiltInStyleParagraphFormatting(int style, LegacyDocParagraphFormat paragraphFormat) {
            OfficeXmlTree styles = _styles;
            int templateProperties = Child(styles, style, "pPr");
            // Imported styles inherit absent controls from their own base style, not from the authoring templates.
            if (templateProperties != OfficeXmlTree.NoNode) {
                foreach (string control in TemplateParagraphControls) RemoveAll(styles, templateProperties, control);
            }

            if (!paragraphFormat.HasFormatting) return;
            int properties = templateProperties != OfficeXmlTree.NoNode ? templateProperties : Append(styles, style, Element(styles, "pPr"));
            string? alignment = ParagraphAlignmentToken(paragraphFormat.Alignment);
            if (alignment != null) ReplaceStyleProperty(properties, Element(styles, "jc", alignment));
            if (paragraphFormat.NumberingListIndex != null) {
                ReplaceStyleProperty(properties, CreateNumberingProperties(styles, paragraphFormat.NumberingListIndex.Value, paragraphFormat.NumberingLevel ?? 0));
            }

            if (paragraphFormat.SpacingBeforeTwips != null || paragraphFormat.SpacingAfterTwips != null || paragraphFormat.LineSpacingTwips != null) {
                int spacing = GetOrAdd(styles, properties, "spacing", ParagraphPropertiesOrder);
                if (paragraphFormat.SpacingBeforeTwips != null) Set(styles, spacing, "before", Text(paragraphFormat.SpacingBeforeTwips.Value));
                if (paragraphFormat.SpacingAfterTwips != null) Set(styles, spacing, "after", Text(paragraphFormat.SpacingAfterTwips.Value));
                ApplyLineSpacing(styles, spacing, paragraphFormat);
            }

            if (paragraphFormat.LeftIndentTwips != null || paragraphFormat.RightIndentTwips != null || paragraphFormat.FirstLineIndentTwips != null) {
                SetIndentation(styles, GetOrAdd(styles, properties, "ind", ParagraphPropertiesOrder), paragraphFormat);
            }

            AddStyleParagraphControls(styles, properties, paragraphFormat, replace: true);
        }

        private void ReplaceStyleProperty(int parent, int replacement) => Replace(_styles, parent, replacement, ParagraphPropertiesOrder);

        /// <summary>Writes the paragraph controls a style can carry, from tabs on, in schema order.</summary>
        private void AddStyleParagraphControls(OfficeXmlTree tree, int properties, LegacyDocParagraphFormat paragraphFormat, bool replace) {
            void Add(int element) {
                if (replace) Replace(tree, properties, element, ParagraphPropertiesOrder);
                else Insert(tree, properties, element, ParagraphPropertiesOrder);
            }

            int tabs = CreateTabs(tree, paragraphFormat.TabStops);
            if (tabs != OfficeXmlTree.NoNode) Add(tabs);
            if (paragraphFormat.KeepLinesTogether.HasValue) Add(OnOffValue(tree, "keepLines", paragraphFormat.KeepLinesTogether.Value));
            if (paragraphFormat.KeepWithNext.HasValue) Add(OnOffValue(tree, "keepNext", paragraphFormat.KeepWithNext.Value));
            if (paragraphFormat.PageBreakBefore.HasValue) Add(OnOffValue(tree, "pageBreakBefore", paragraphFormat.PageBreakBefore.Value));
            if (paragraphFormat.AvoidWidowAndOrphan.HasValue) Add(OnOffValue(tree, "widowControl", paragraphFormat.AvoidWidowAndOrphan.Value));
            if (paragraphFormat.SuppressLineNumbers.HasValue) Add(OnOffValue(tree, "suppressLineNumbers", paragraphFormat.SuppressLineNumbers.Value));
            if (paragraphFormat.SuppressAutoHyphens.HasValue) Add(OnOffValue(tree, "suppressAutoHyphens", paragraphFormat.SuppressAutoHyphens.Value));
            if (paragraphFormat.ContextualSpacing.HasValue) Add(OnOffValue(tree, "contextualSpacing", paragraphFormat.ContextualSpacing.Value));
            if (paragraphFormat.MirrorIndents.HasValue) Add(OnOffValue(tree, "mirrorIndents", paragraphFormat.MirrorIndents.Value));
            if (paragraphFormat.Kinsoku.HasValue) Add(OnOffValue(tree, "kinsoku", paragraphFormat.Kinsoku.Value));
            if (paragraphFormat.WordWrap.HasValue) Add(OnOffValue(tree, "wordWrap", paragraphFormat.WordWrap.Value));
            if (paragraphFormat.OverflowPunctuation.HasValue) Add(OnOffValue(tree, "overflowPunct", paragraphFormat.OverflowPunctuation.Value));
            if (paragraphFormat.TopLinePunctuation.HasValue) Add(OnOffValue(tree, "topLinePunct", paragraphFormat.TopLinePunctuation.Value));
            if (paragraphFormat.AutoSpaceDE.HasValue) Add(OnOffValue(tree, "autoSpaceDE", paragraphFormat.AutoSpaceDE.Value));
            if (paragraphFormat.AutoSpaceDN.HasValue) Add(OnOffValue(tree, "autoSpaceDN", paragraphFormat.AutoSpaceDN.Value));
            if (paragraphFormat.Bidirectional.HasValue) Add(OnOffValue(tree, "bidi", paragraphFormat.Bidirectional.Value));
            string? verticalAlignment = VerticalCharacterAlignmentToken(paragraphFormat.VerticalCharacterAlignment);
            if (verticalAlignment != null) Add(Element(tree, "textAlignment", verticalAlignment));
            if (paragraphFormat.OutlineLevel != null) Add(Element(tree, "outlineLvl", Text(paragraphFormat.OutlineLevel.Value)));
            if (paragraphFormat.ParagraphShading != null && !string.IsNullOrEmpty(paragraphFormat.ParagraphShading.Value.FillColorHex)) {
                int shading = Element(tree, "shd");
                Set(tree, shading, "val", "clear");
                Set(tree, shading, "color", "auto");
                Set(tree, shading, "fill", paragraphFormat.ParagraphShading.Value.OpenXmlFill!);
                Add(shading);
            }

            if (paragraphFormat.ParagraphBorders != null && paragraphFormat.ParagraphBorders.Value.HasAny) {
                Add(CreateParagraphBorders(tree, paragraphFormat.ParagraphBorders.Value));
            }
        }

        /// <summary>The paragraph properties of a style, list level or other definition, or NoNode when there are none.</summary>
        private int CreateStyleParagraphProperties(OfficeXmlTree tree, LegacyDocParagraphFormat paragraphFormat, string elementName = "pPr") {
            int properties = Element(tree, elementName);
            string? alignment = ParagraphAlignmentToken(paragraphFormat.Alignment);
            if (alignment != null) Insert(tree, properties, Element(tree, "jc", alignment), ParagraphPropertiesOrder);
            if (paragraphFormat.NumberingListIndex != null) {
                Insert(tree, properties, CreateNumberingProperties(tree, paragraphFormat.NumberingListIndex.Value, paragraphFormat.NumberingLevel ?? 0), ParagraphPropertiesOrder);
            }

            if (paragraphFormat.SpacingBeforeTwips != null || paragraphFormat.SpacingAfterTwips != null || paragraphFormat.LineSpacingTwips != null) {
                int spacing = Element(tree, "spacing");
                if (paragraphFormat.SpacingBeforeTwips != null) Set(tree, spacing, "before", Text(paragraphFormat.SpacingBeforeTwips.Value));
                if (paragraphFormat.SpacingAfterTwips != null) Set(tree, spacing, "after", Text(paragraphFormat.SpacingAfterTwips.Value));
                ApplyLineSpacing(tree, spacing, paragraphFormat);
                Insert(tree, properties, spacing, ParagraphPropertiesOrder);
            }

            if (paragraphFormat.LeftIndentTwips != null || paragraphFormat.RightIndentTwips != null || paragraphFormat.FirstLineIndentTwips != null) {
                int indentation = Element(tree, "ind");
                SetIndentation(tree, indentation, paragraphFormat);
                Insert(tree, properties, indentation, ParagraphPropertiesOrder);
            }

            AddStyleParagraphControls(tree, properties, paragraphFormat, replace: false);
            return HasElements(tree, properties) ? properties : OfficeXmlTree.NoNode;
        }

        private void MergeBuiltInStyleRunFormatting(int style, LegacyDocCharacterFormat characterFormat, LegacyDocCharacterFormatProperties styleToggles) {
            OfficeXmlTree styles = _styles;
            int templateProperties = Child(styles, style, "rPr");
            // Missing source formatting inherits through basedOn. Template sizes, fonts, colors and effects must not
            // override it.
            if (templateProperties != OfficeXmlTree.NoNode) RemoveChildren(styles, templateProperties);
            if (!characterFormat.HasFormatting) return;
            int properties = templateProperties != OfficeXmlTree.NoNode ? templateProperties : Append(styles, style, Element(styles, "rPr"));
            int created = CreateStyleRunProperties(styles, characterFormat, styleToggles);
            if (created != OfficeXmlTree.NoNode) {
                foreach (int child in Elements(styles, created)) {
                    styles.Remove(child);
                    Insert(styles, properties, child, RunPropertiesOrder);
                }
            }
        }

        // ---- Numbering ------------------------------------------------------------------------------------------

        private OfficeXmlTree Numbering() {
            if (_numbering != null) return _numbering;
            OfficeXmlTree tree = _package.CreateXmlPart(NumberingPartName, NumberingContentType, W, "numbering");
            tree.DeclareNamespace(tree.Root, "w", OfficeXmlNamespaces.KnownUri(W)!);
            _package.AddRelationship(MainPartName, RelationshipTypePrefix + "numbering", NumberingPartName, id: NextRelationshipId());
            return _numbering = tree;
        }

        private void AddNumberingDefinitions(LegacyDocNumbering source) {
            if (source.Instances.Count == 0) return;
            OfficeXmlTree numbering = Numbering();
            var ids = new Dictionary<int, int>();
            var definitions = source.Definitions.ToDictionary(item => item.Id);
            foreach (LegacyDocListDefinition definition in source.Definitions) {
                int id = ids.Count;
                ids.Add(definition.Id, id);
                int abstractNum = Element(numbering, "abstractNum");
                Set(numbering, abstractNum, "abstractNumId", Text(id));
                Append(numbering, abstractNum, Element(numbering, "multiLevelType", definition.Levels.Count == 1 ? "singleLevel"
                    : definition.Hybrid ? "hybridMultilevel" : "multilevel"));
                for (int level = 0; level < definition.Levels.Count; level++) {
                    Append(numbering, abstractNum, CreateNumberingLevel(numbering, definition.Levels[level], level));
                }

                Append(numbering, numbering.Root, abstractNum);
            }

            for (int index = 0; index < source.Instances.Count; index++) {
                LegacyDocListInstance sourceInstance = source.Instances[index];
                int instance = Element(numbering, "num");
                Set(numbering, instance, "numId", Text(index + 1));
                Append(numbering, instance, Element(numbering, "abstractNumId", Text(ids[sourceInstance.DefinitionId])));
                LegacyDocListDefinition definition = definitions[sourceInstance.DefinitionId];
                foreach (LegacyDocListOverride item in sourceInstance.Overrides) {
                    int levelOverride = Element(numbering, "lvlOverride");
                    Set(numbering, levelOverride, "ilvl", Text(item.Level));
                    if (item.Start.HasValue) Append(numbering, levelOverride, Element(numbering, "startOverride", Text(item.Start.Value)));
                    if (item.Formatting != null) {
                        int level = CreateNumberingLevel(numbering, item.Formatting, item.Level);
                        // LFOLVL.fFormatting alone does not override the abstract start.
                        if (!item.Start.HasValue && item.Level < definition.Levels.Count) {
                            Set(numbering, Child(numbering, level, "start"), "val", Text(definition.Levels[item.Level].Start));
                        }

                        Append(numbering, levelOverride, level);
                    }

                    Append(numbering, instance, levelOverride);
                }

                Append(numbering, numbering.Root, instance);
            }
        }

        private int CreateNumberingLevel(OfficeXmlTree numbering, LegacyDocListLevel source, int index) {
            int level = Element(numbering, "lvl");
            Set(numbering, level, "ilvl", Text(index));
            Set(numbering, level, "tentative", Text((source.Flags & 128) != 0));
            Insert(numbering, level, Element(numbering, "start", Text(source.Start)), LevelOrder);
            Insert(numbering, level, Element(numbering, "numFmt", source.Format.ToXml()), LevelOrder);
            if ((source.Flags & 8) != 0) Insert(numbering, level, Element(numbering, "lvlRestart", Text(source.RestartLimit)), LevelOrder);
            if (source.StyleIndex.HasValue && _styleSheet.TryGetParagraphStyle(source.StyleIndex.Value, out LegacyDocParagraphStyle style)) {
                string? styleId = style.BuiltInStyle?.ToStringStyle() ?? style.StyleId;
                if (!string.IsNullOrEmpty(styleId)) Insert(numbering, level, Element(numbering, "pStyle", styleId!), LevelOrder);
            }

            if ((source.Flags & 4) != 0) Insert(numbering, level, Element(numbering, "isLgl"), LevelOrder);
            Insert(numbering, level, Element(numbering, "suff", source.Follow == 0 ? "tab" : source.Follow == 1 ? "space" : "nothing"), LevelOrder);
            Insert(numbering, level, Element(numbering, "lvlText", source.Text), LevelOrder);
            Insert(numbering, level, Element(numbering, "lvlJc", (source.Flags & 3) == 2 ? "right" : (source.Flags & 3) == 1 ? "center" : "left"), LevelOrder);
            int paragraph = CreateStyleParagraphProperties(numbering, source.ParagraphFormat);
            if (paragraph != OfficeXmlTree.NoNode) Insert(numbering, level, paragraph, LevelOrder);
            int run = CreateStyleRunProperties(numbering, source.CharacterFormat);
            if (run != OfficeXmlTree.NoNode) Insert(numbering, level, run, LevelOrder);
            return level;
        }

        private int CreateNumberingProperties(OfficeXmlTree tree, ushort listIndex, byte level) {
            EnsureNumberingDefinition(listIndex);
            int numbering = Element(tree, "numPr");
            ReplaceNumberingProperties(tree, numbering, listIndex, level);
            return numbering;
        }

        private static void ReplaceNumberingProperties(OfficeXmlTree tree, int numberingProperties, ushort listIndex, byte level) {
            RemoveAll(tree, numberingProperties, "ilvl");
            RemoveAll(tree, numberingProperties, "numId");
            Append(tree, numberingProperties, Element(tree, "ilvl", Text(level)));
            Append(tree, numberingProperties, Element(tree, "numId", Text(listIndex)));
        }

        /// <summary>Adds a decimal list for a list reference the source numbering doesn't define.</summary>
        private void EnsureNumberingDefinition(int numberId) {
            OfficeXmlTree numbering = Numbering();
            int nextAbstractId = 0;
            foreach (int instance in Elements(numbering, numbering.Root, "num")) {
                if (Get(numbering, instance, "numId") == Text(numberId)) return;
            }

            foreach (int abstractNum in Elements(numbering, numbering.Root, "abstractNum")) {
                if (int.TryParse(Get(numbering, abstractNum, "abstractNumId"), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int id)) {
                    nextAbstractId = Math.Max(nextAbstractId, id + 1);
                }
            }

            int fallback = Element(numbering, "abstractNum");
            Set(numbering, fallback, "abstractNumId", Text(nextAbstractId));
            Append(numbering, fallback, Element(numbering, "multiLevelType", "hybridMultilevel"));
            for (int level = 0; level <= 8; level++) {
                int numberingLevel = Element(numbering, "lvl");
                Set(numbering, numberingLevel, "ilvl", Text(level));
                Set(numbering, numberingLevel, "tentative", Text(level > 0));
                Append(numbering, numberingLevel, Element(numbering, "start", "1"));
                Append(numbering, numberingLevel, Element(numbering, "numFmt", (level % 3) switch { 1 => "lowerLetter", 2 => "lowerRoman", _ => "decimal" }));
                Append(numbering, numberingLevel, Element(numbering, "lvlText", "%" + Text(level + 1) + "."));
                Append(numbering, numberingLevel, Element(numbering, "lvlJc", "left"));
                int properties = Append(numbering, numberingLevel, Element(numbering, "pPr"));
                int indentation = Append(numbering, properties, Element(numbering, "ind"));
                Set(numbering, indentation, "left", Text((level + 1) * 720));
                Set(numbering, indentation, "hanging", "360");
                Append(numbering, fallback, numberingLevel);
            }

            Append(numbering, numbering.Root, fallback);
            int numberingInstance = Element(numbering, "num");
            Set(numbering, numberingInstance, "numId", Text(numberId));
            Append(numbering, numberingInstance, Element(numbering, "abstractNumId", Text(nextAbstractId)));
            Append(numbering, numbering.Root, numberingInstance);
        }

        /// <summary>A relationship id in the form the Open XML SDK generates for parts and hyperlinks it adds.</summary>
        private static string NextRelationshipId() => "R" + Guid.NewGuid().ToString("N").Substring(0, 16);
    }
}
