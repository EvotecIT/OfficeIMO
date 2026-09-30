namespace OfficeIMO.Rtf;

public sealed partial class RtfDocument {
    /// <summary>Resolves the list instance, definition and targeted level beneath a paragraph. The returned level is an independent copy.</summary>
    public RtfListFormatting? ResolveListFormatting(RtfParagraph paragraph) {
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
        return ResolveListFormattingCore(ApplyParagraphInheritance(paragraph.CopyFormattingView()));
    }

    private RtfListFormatting? ResolveListFormattingCore(RtfParagraph paragraph) {
        if (paragraph.ListId == 0) return null;
        RtfListOverride? instance = paragraph.ListId.HasValue
            ? ListOverrides.LastOrDefault(item => item.Id == paragraph.ListId.Value) : null;
        int? definitionId = instance?.ListId ?? paragraph.ListDefinitionId;
        RtfListDefinition? definition = definitionId.HasValue
            ? ListDefinitions.LastOrDefault(item => item.Id == definitionId.Value) : null;
        int levelIndex = Math.Max(0, paragraph.ListLevel ?? 0);
        RtfListLevel? source = definition?.Levels.LastOrDefault(item => item.LevelIndex == levelIndex);
        int? originalStart = source?.StartAt;
        RtfListLevelOverride? levelOverride = instance?.LevelOverrides.Select((item, index) => new { Item = item, Index = index })
            .LastOrDefault(item => (item.Item.LevelIndex ?? item.Index) == levelIndex)?.Item;
        if (levelOverride?.OverrideFormat == true && levelOverride.Formatting != null) source = levelOverride.Formatting;
        if (source == null && paragraph.ListKind == RtfListKind.None && !paragraph.LegacyNumbering.HasAnyValue) return null;
        RtfListLevel level = source == null
            ? new RtfListLevel(levelIndex, paragraph.ListKind == RtfListKind.None ? RtfListKind.Decimal : paragraph.ListKind)
            : new RtfCloneContext().Clone(source)!;
        if (source == null && paragraph.LegacyNumbering.HasAnyValue) {
            RtfLegacyNumbering legacy = paragraph.LegacyNumbering;
            level.NumberFormat = legacy.NumberStyle switch {
                RtfLegacyNumberingStyle.UpperRoman => 1, RtfLegacyNumberingStyle.LowerRoman => 2,
                RtfLegacyNumberingStyle.UpperLetter => 3, RtfLegacyNumberingStyle.LowerLetter => 4,
                RtfLegacyNumberingStyle.Ordinal => 5, RtfLegacyNumberingStyle.Cardinal => 6,
                RtfLegacyNumberingStyle.OrdinalText => 7, _ => 0
            };
            if (legacy.TextBefore != null || legacy.TextAfter != null) {
                level.Text = (legacy.TextBefore ?? string.Empty) + (level.Kind == RtfListKind.Bullet ? string.Empty : "%" + (levelIndex + 1).ToString(CultureInfo.InvariantCulture)) + (legacy.TextAfter ?? string.Empty);
            }
            level.LeftIndentTwips = legacy.IndentTwips;
        }
        if (levelOverride?.OverrideFormat == true && levelOverride.Formatting != null) level.StartAt = originalStart;
        if (levelOverride?.OverrideStartAt == true) level.StartAt = levelOverride.StartAt ?? levelOverride.Formatting?.StartAt ?? level.StartAt;
        level.StartAt ??= paragraph.LegacyNumbering.StartAt ?? 1;
        return new RtfListFormatting(paragraph.ListId, definitionId, levelIndex, level, source != null);
    }
}

/// <summary>Effective numbering for one paragraph, including its distinct list-instance identity.</summary>
public sealed class RtfListFormatting {
    internal RtfListFormatting(int? instanceId, int? definitionId, int levelIndex, RtfListLevel level, bool isDefined) {
        InstanceId = instanceId;
        DefinitionId = definitionId;
        LevelIndex = levelIndex;
        Level = level;
        IsDefined = isDefined;
    }

    /// <summary>List override id used by this paragraph.</summary>
    public int? InstanceId { get; }
    /// <summary>Referenced definition id.</summary>
    public int? DefinitionId { get; }
    /// <summary>Zero-based target level.</summary>
    public int LevelIndex { get; }
    /// <summary>Independent effective level formatting.</summary>
    public RtfListLevel Level { get; }
    /// <summary>Whether the paragraph resolves to an authored definition or formatting override.</summary>
    public bool IsDefined { get; }
    /// <summary>Counter identity. Instances that share a definition retain independent numbering.</summary>
    public string Identity => InstanceId.HasValue ? "instance:" + InstanceId.Value.ToString(CultureInfo.InvariantCulture)
        : DefinitionId.HasValue ? "definition:" + DefinitionId.Value.ToString(CultureInfo.InvariantCulture) : "anonymous:" + Level.Kind;
}
