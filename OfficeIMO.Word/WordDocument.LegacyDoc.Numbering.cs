using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private static void AddLegacyDocNumberingDefinitions(WordDocument document, LegacyDocNumbering source, LegacyDocStyleSheet styleSheet) {
            if (source.Instances.Count == 0) return;
            MainDocumentPart main = document.MainDocumentPartRoot;
            NumberingDefinitionsPart part = main.NumberingDefinitionsPart ?? main.AddNewPart<NumberingDefinitionsPart>();
            Numbering numbering = part.Numbering ??= new Numbering();
            var ids = new Dictionary<int, int>();
            var definitions = source.Definitions.ToDictionary(item => item.Id);
            foreach (LegacyDocListDefinition definition in source.Definitions) {
                int id = ids.Count;
                ids.Add(definition.Id, id);
                var abstractNum = new AbstractNum { AbstractNumberId = id };
                abstractNum.Append(new MultiLevelType { Val = definition.Levels.Count == 1 ? MultiLevelValues.SingleLevel
                    : definition.Hybrid ? MultiLevelValues.HybridMultilevel : MultiLevelValues.Multilevel });
                for (int level = 0; level < definition.Levels.Count; level++)
                    abstractNum.Append(CreateLegacyDocNumberingLevel(document, definition.Levels[level], level, styleSheet));
                numbering.Append(abstractNum);
            }
            for (int index = 0; index < source.Instances.Count; index++) {
                LegacyDocListInstance sourceInstance = source.Instances[index];
                var instance = new NumberingInstance(new AbstractNumId { Val = ids[sourceInstance.DefinitionId] }) { NumberID = index + 1 };
                LegacyDocListDefinition definition = definitions[sourceInstance.DefinitionId];
                foreach (LegacyDocListOverride item in sourceInstance.Overrides) {
                    var levelOverride = new LevelOverride { LevelIndex = item.Level };
                    if (item.Start.HasValue) levelOverride.Append(new StartOverrideNumberingValue { Val = item.Start.Value });
                    if (item.Formatting != null) {
                        Level level = CreateLegacyDocNumberingLevel(document, item.Formatting, item.Level, styleSheet);
                        // LFOLVL.fFormatting alone does not override the abstract start.
                        if (!item.Start.HasValue && item.Level < definition.Levels.Count)
                            level.StartNumberingValue!.Val = definition.Levels[item.Level].Start;
                        levelOverride.Append(level);
                    }
                    instance.Append(levelOverride);
                }
                numbering.Append(instance);
            }
        }

        private static Level CreateLegacyDocNumberingLevel(WordDocument document, LegacyDocListLevel source, int index, LegacyDocStyleSheet styles) {
            var level = new Level { LevelIndex = index, Tentative = (source.Flags & 128) != 0 };
            level.AddChild(new StartNumberingValue { Val = source.Start }, true);
            level.AddChild(new NumberingFormat { Val = source.Format }, true);
            if ((source.Flags & 8) != 0) level.AddChild(new LevelRestart { Val = source.RestartLimit }, true);
            if (source.StyleIndex.HasValue && styles.TryGetParagraphStyle(source.StyleIndex.Value, out LegacyDocParagraphStyle style)) {
                string? styleId = style.BuiltInStyle?.ToStringStyle() ?? style.StyleId;
                if (!string.IsNullOrEmpty(styleId)) level.AddChild(new ParagraphStyleIdInLevel { Val = styleId }, true);
            }
            if ((source.Flags & 4) != 0) level.AddChild(new IsLegalNumberingStyle(), true);
            level.AddChild(new LevelSuffix { Val = source.Follow == 0 ? LevelSuffixValues.Tab : source.Follow == 1 ? LevelSuffixValues.Space : LevelSuffixValues.Nothing }, true);
            level.AddChild(new LevelText { Val = source.Text }, true);
            level.AddChild(new LevelJustification { Val = (source.Flags & 3) == 2 ? LevelJustificationValues.Right
                : (source.Flags & 3) == 1 ? LevelJustificationValues.Center : LevelJustificationValues.Left }, true);
            StyleParagraphProperties? paragraph = CreateLegacyDocStyleParagraphProperties(document, source.ParagraphFormat);
            if (paragraph != null) {
                var properties = new PreviousParagraphProperties();
                foreach (OpenXmlElement child in paragraph.ChildElements) properties.AddChild(child.CloneNode(true), true);
                level.AddChild(properties, true);
            }
            StyleRunProperties? run = CreateLegacyDocStyleRunProperties(source.CharacterFormat);
            if (run != null) {
                var properties = new NumberingSymbolRunProperties();
                foreach (OpenXmlElement child in run.ChildElements) properties.AddChild(child.CloneNode(true), true);
                level.AddChild(properties, true);
            }
            return level;
        }
    }
}
