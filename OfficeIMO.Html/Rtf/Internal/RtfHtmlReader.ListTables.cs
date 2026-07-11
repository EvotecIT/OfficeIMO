using System.Globalization;

namespace OfficeIMO.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void ApplyListTables(Dictionary<string, string> values) {
            var definitions = new List<RtfListDefinition>();
            for (int index = 0; ; index++) {
                string prefix = "definition." + index.ToString(CultureInfo.InvariantCulture);
                int? id = ReadInt(values, prefix + ".id");
                if (!id.HasValue) {
                    break;
                }

                var definition = new RtfListDefinition(id.Value) {
                    TemplateId = ReadInt(values, prefix + ".templateId"),
                    Name = ReadString(values, prefix + ".name")
                };

                int levelCount = ReadInt(values, prefix + ".level.count") ?? 0;
                for (int levelIndex = 0; levelIndex < levelCount; levelIndex++) {
                    AddListLevel(definition, values, prefix + ".level." + levelIndex.ToString(CultureInfo.InvariantCulture));
                }

                definitions.Add(definition);
            }

            var overrides = new List<RtfListOverride>();
            for (int index = 0; ; index++) {
                string prefix = "override." + index.ToString(CultureInfo.InvariantCulture);
                int? id = ReadInt(values, prefix + ".id");
                int? listId = ReadInt(values, prefix + ".listId");
                if (!id.HasValue || !listId.HasValue) {
                    break;
                }

                var listOverride = new RtfListOverride(id.Value, listId.Value) {
                    OverrideCount = ReadInt(values, prefix + ".overrideCount")
                };

                int levelOverrideCount = ReadInt(values, prefix + ".levelOverride.count") ?? 0;
                for (int levelOverrideIndex = 0; levelOverrideIndex < levelOverrideCount; levelOverrideIndex++) {
                    AddListLevelOverride(listOverride, values, prefix + ".levelOverride." + levelOverrideIndex.ToString(CultureInfo.InvariantCulture));
                }

                overrides.Add(listOverride);
            }

            if (definitions.Count > 0) {
                _document.ReplaceListDefinitions(definitions);
            }

            if (overrides.Count > 0) {
                _document.ReplaceListOverrides(overrides);
            }
        }

        private static void AddListLevel(RtfListDefinition definition, Dictionary<string, string> values, string prefix) {
            RtfListLevel level = definition.AddLevel(ReadEnum(values, prefix + ".kind", RtfListKind.Decimal));
            level.NumberFormat = ReadInt(values, prefix + ".numberFormat");
            level.NumberFormatN = ReadInt(values, prefix + ".numberFormatN");
            level.Alignment = ReadEnum<RtfListLevelAlignment>(values, prefix + ".alignment");
            level.AlignmentN = ReadEnum<RtfListLevelAlignment>(values, prefix + ".alignmentN");
            level.FollowCharacter = ReadEnum<RtfListLevelFollowCharacter>(values, prefix + ".followCharacter");
            level.StartAt = ReadInt(values, prefix + ".startAt");
            level.SpaceTwips = ReadInt(values, prefix + ".spaceTwips");
            level.IndentTwips = ReadInt(values, prefix + ".indentTwips");
            level.LegalNumbering = ReadBool(values, prefix + ".legalNumbering");
            level.NoRestart = ReadBool(values, prefix + ".noRestart");
            level.PictureIndex = ReadInt(values, prefix + ".pictureIndex");
            level.PictureNoSize = ReadBool(values, prefix + ".pictureNoSize") == true;
            level.Text = ReadString(values, prefix + ".text");
            level.Numbers = ReadString(values, prefix + ".numbers");
            level.LeftIndentTwips = ReadInt(values, prefix + ".leftIndentTwips");
            level.FirstLineIndentTwips = ReadInt(values, prefix + ".firstLineIndentTwips");
        }

        private static void AddListLevelOverride(RtfListOverride listOverride, Dictionary<string, string> values, string prefix) {
            RtfListLevelOverride levelOverride = listOverride.AddLevelOverride();
            levelOverride.LevelIndex = ReadInt(values, prefix + ".levelIndex") ?? levelOverride.LevelIndex;
            levelOverride.OverrideFormat = ReadBool(values, prefix + ".overrideFormat");
            levelOverride.OverrideStartAt = ReadBool(values, prefix + ".overrideStartAt");
            levelOverride.StartAt = ReadInt(values, prefix + ".startAt");
        }
    }
}
