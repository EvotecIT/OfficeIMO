using System.Globalization;

namespace OfficeIMO.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendListTableMetadata(StringBuilder builder, RtfDocument document, string newline) {
        if (document.ListDefinitions.Count == 0 && document.ListOverrides.Count == 0) {
            return;
        }

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int index = 0; index < document.ListDefinitions.Count; index++) {
            AddListDefinition(values, "definition." + index.ToString(CultureInfo.InvariantCulture), document.ListDefinitions[index]);
        }

        for (int index = 0; index < document.ListOverrides.Count; index++) {
            AddListOverride(values, "override." + index.ToString(CultureInfo.InvariantCulture), document.ListOverrides[index]);
        }

        builder.Append(newline);
        builder.Append("<meta name=\"officeimo-rtf-lists\" content=\"");
        builder.Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values)));
        builder.Append("\">");
    }

    private static void AddListDefinition(Dictionary<string, string> values, string prefix, RtfListDefinition definition) {
        AddInt(values, prefix + ".id", definition.Id);
        AddNullableInt(values, prefix + ".templateId", definition.TemplateId);
        AddString(values, prefix + ".name", definition.Name);
        AddInt(values, prefix + ".level.count", definition.Levels.Count);
        for (int index = 0; index < definition.Levels.Count; index++) {
            AddListLevel(values, prefix + ".level." + index.ToString(CultureInfo.InvariantCulture), definition.Levels[index]);
        }
    }

    private static void AddListLevel(Dictionary<string, string> values, string prefix, RtfListLevel level) {
        AddEnum(values, prefix + ".kind", (RtfListKind?)level.Kind);
        AddNullableInt(values, prefix + ".numberFormat", level.NumberFormat);
        AddNullableInt(values, prefix + ".numberFormatN", level.NumberFormatN);
        AddEnum(values, prefix + ".alignment", level.Alignment);
        AddEnum(values, prefix + ".alignmentN", level.AlignmentN);
        AddEnum(values, prefix + ".followCharacter", level.FollowCharacter);
        AddNullableInt(values, prefix + ".startAt", level.StartAt);
        AddNullableInt(values, prefix + ".spaceTwips", level.SpaceTwips);
        AddNullableInt(values, prefix + ".indentTwips", level.IndentTwips);
        AddNullableBool(values, prefix + ".legalNumbering", level.LegalNumbering);
        AddNullableBool(values, prefix + ".noRestart", level.NoRestart);
        AddNullableInt(values, prefix + ".pictureIndex", level.PictureIndex);
        AddBool(values, prefix + ".pictureNoSize", level.PictureNoSize);
        AddString(values, prefix + ".text", level.Text);
        AddString(values, prefix + ".numbers", level.Numbers);
        AddNullableInt(values, prefix + ".leftIndentTwips", level.LeftIndentTwips);
        AddNullableInt(values, prefix + ".firstLineIndentTwips", level.FirstLineIndentTwips);
    }

    private static void AddListOverride(Dictionary<string, string> values, string prefix, RtfListOverride listOverride) {
        AddInt(values, prefix + ".id", listOverride.Id);
        AddInt(values, prefix + ".listId", listOverride.ListId);
        AddNullableInt(values, prefix + ".overrideCount", listOverride.OverrideCount);
        AddInt(values, prefix + ".levelOverride.count", listOverride.LevelOverrides.Count);
        for (int index = 0; index < listOverride.LevelOverrides.Count; index++) {
            AddListLevelOverride(values, prefix + ".levelOverride." + index.ToString(CultureInfo.InvariantCulture), listOverride.LevelOverrides[index]);
        }
    }

    private static void AddListLevelOverride(Dictionary<string, string> values, string prefix, RtfListLevelOverride levelOverride) {
        AddNullableBool(values, prefix + ".overrideFormat", levelOverride.OverrideFormat);
        AddNullableBool(values, prefix + ".overrideStartAt", levelOverride.OverrideStartAt);
        AddNullableInt(values, prefix + ".startAt", levelOverride.StartAt);
    }
}
