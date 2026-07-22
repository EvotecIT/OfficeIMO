namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteListTables(StringBuilder builder, RtfDocument document, int unicodeSkipCount) {
        EffectiveListTables lists = BuildEffectiveListTables(document);
        if (lists.Definitions.Count == 0 && lists.Overrides.Count == 0) return;

        builder.Append(@"{\*\listtable");
        foreach (RtfListDefinition definition in lists.Definitions.OrderBy(definition => definition.Id)) {
            builder.Append(@"{\list");
            if (definition.TemplateId.HasValue) {
                builder.Append(@"\listtemplateid");
                builder.Append(definition.TemplateId.Value.ToString(CultureInfo.InvariantCulture));
            }

            foreach (RtfListLevel level in definition.Levels.OrderBy(level => level.LevelIndex)) {
                WriteListLevel(builder, level, unicodeSkipCount);
            }

            builder.Append(@"{\listname ");
            builder.Append(EscapeText(definition.Name ?? string.Empty, unicodeSkipCount));
            builder.Append(";}");
            builder.Append(@"\listid");
            builder.Append(definition.Id.ToString(CultureInfo.InvariantCulture));
            builder.Append('}');
        }

        builder.Append('}');

        builder.Append(@"{\*\listoverridetable");
        foreach (RtfListOverride listOverride in lists.Overrides.OrderBy(listOverride => listOverride.Id)) {
            builder.Append(@"{\listoverride\listid");
            builder.Append(listOverride.ListId.ToString(CultureInfo.InvariantCulture));
            builder.Append(@"\listoverridecount");
            builder.Append((listOverride.OverrideCount ?? listOverride.LevelOverrides.Count).ToString(CultureInfo.InvariantCulture));
            foreach (RtfListLevelOverride levelOverride in listOverride.LevelOverrides) {
                WriteListLevelOverride(builder, levelOverride);
            }

            builder.Append(@"\ls");
            builder.Append(listOverride.Id.ToString(CultureInfo.InvariantCulture));
            builder.Append('}');
        }

        builder.Append('}');
    }

    private static void WriteListLevelOverride(StringBuilder builder, RtfListLevelOverride levelOverride) {
        if (!levelOverride.HasAnyValue) return;

        builder.Append(@"{\lfolevel");
        AppendOptionalBinary(builder, @"\listoverrideformat", levelOverride.OverrideFormat);
        AppendOptionalBinary(builder, @"\listoverridestartat", levelOverride.OverrideStartAt);
        AppendOptionalTwips(builder, @"\levelstartat", levelOverride.StartAt);
        builder.Append('}');
    }

    private static void WriteListLevel(StringBuilder builder, RtfListLevel level, int unicodeSkipCount) {
        int numberFormat = level.NumberFormat ?? level.NumberFormatN ?? (level.Kind == RtfListKind.Bullet ? 23 : 0);
        int numberFormatN = level.NumberFormatN ?? numberFormat;
        builder.Append(@"{\listlevel\levelnfc");
        builder.Append(numberFormat.ToString(CultureInfo.InvariantCulture));
        builder.Append(@"\levelnfcn");
        builder.Append(numberFormatN.ToString(CultureInfo.InvariantCulture));
        builder.Append(@"\leveljc");
        builder.Append(ToRtfListLevelAlignmentValue(level.Alignment).ToString(CultureInfo.InvariantCulture));
        builder.Append(@"\leveljcn");
        builder.Append(ToRtfListLevelAlignmentValue(level.AlignmentN ?? level.Alignment).ToString(CultureInfo.InvariantCulture));
        builder.Append(@"\levelfollow");
        builder.Append(ToRtfListLevelFollowValue(level.FollowCharacter).ToString(CultureInfo.InvariantCulture));
        builder.Append(@"\levelstartat");
        builder.Append((level.StartAt ?? 1).ToString(CultureInfo.InvariantCulture));
        builder.Append(@"\levelspace");
        builder.Append(level.SpaceTwips.GetValueOrDefault().ToString(CultureInfo.InvariantCulture));
        builder.Append(@"\levelindent");
        builder.Append(level.IndentTwips.GetValueOrDefault().ToString(CultureInfo.InvariantCulture));
        AppendOptionalBinary(builder, @"\levellegal", level.LegalNumbering);
        AppendOptionalBinary(builder, @"\levelnorestart", level.NoRestart);
        AppendOptionalTwips(builder, @"\levelpicture", level.PictureIndex);
        if (level.PictureNoSize) {
            builder.Append(@"\levelpicturenosize");
        }

        string levelText = level.Text ?? (level.Kind == RtfListKind.Bullet ? "\u2022" : "%1.");
        builder.Append(@"{\leveltext");
        WriteListText(builder, levelText, unicodeSkipCount);
        builder.Append(";}");
        builder.Append(@"{\levelnumbers");
        WriteListText(builder, level.Numbers ?? (level.Kind == RtfListKind.Bullet ? string.Empty : "\u0001"), unicodeSkipCount);
        builder.Append(";}");
        AppendOptionalTwips(builder, @"\fi", level.FirstLineIndentTwips);
        AppendOptionalTwips(builder, @"\li", level.LeftIndentTwips);
        builder.Append('}');
    }

    private static void WriteListText(StringBuilder builder, string text, int unicodeSkipCount) {
        int count = text.Length;
        builder.Append(@"\'");
        builder.Append(Math.Min(count, 255).ToString("x2", CultureInfo.InvariantCulture));
        builder.Append(EscapeText(text, unicodeSkipCount));
    }

    internal static EffectiveListTables BuildEffectiveListTables(RtfDocument document) {
        var definitions = document.ListDefinitions.ToDictionary(definition => definition.Id, CloneListDefinition);
        var overrides = document.ListOverrides.ToDictionary(listOverride => listOverride.Id, CloneListOverride);

        foreach (RtfParagraph paragraph in EnumerateParagraphs(document)) {
            if (!paragraph.ListId.HasValue || paragraph.ListKind == RtfListKind.None) {
                continue;
            }

            int overrideId = paragraph.ListId.Value;
            if (!overrides.TryGetValue(overrideId, out RtfListOverride? listOverride)) {
                listOverride = new RtfListOverride(overrideId, paragraph.ListDefinitionId ?? overrideId) {
                    OverrideCount = 0
                };
                overrides.Add(overrideId, listOverride);
            }

            if (!definitions.TryGetValue(listOverride.ListId, out RtfListDefinition? definition)) {
                definition = new RtfListDefinition(listOverride.ListId) {
                    Name = paragraph.ListKind == RtfListKind.Bullet ? "Bullet" : "Numbered"
                };
                definitions.Add(definition.Id, definition);
            }

            EnsureListLevel(definition, paragraph);
        }

        return new EffectiveListTables(definitions.Values.ToList(), overrides.Values.ToList());
    }

    private static RtfListDefinition CloneListDefinition(RtfListDefinition source) {
        var definition = new RtfListDefinition(source.Id) {
            TemplateId = source.TemplateId,
            Name = source.Name
        };
        foreach (RtfListLevel level in source.Levels) {
            definition.AddParsedLevel(new RtfListLevel(level.LevelIndex, level.Kind) {
                NumberFormat = level.NumberFormat,
                NumberFormatN = level.NumberFormatN,
                Alignment = level.Alignment,
                AlignmentN = level.AlignmentN,
                FollowCharacter = level.FollowCharacter,
                StartAt = level.StartAt,
                SpaceTwips = level.SpaceTwips,
                IndentTwips = level.IndentTwips,
                LegalNumbering = level.LegalNumbering,
                NoRestart = level.NoRestart,
                PictureIndex = level.PictureIndex,
                PictureNoSize = level.PictureNoSize,
                Text = level.Text,
                Numbers = level.Numbers,
                LeftIndentTwips = level.LeftIndentTwips,
                FirstLineIndentTwips = level.FirstLineIndentTwips
            });
        }

        return definition;
    }

    private static RtfListOverride CloneListOverride(RtfListOverride source) {
        var listOverride = new RtfListOverride(source.Id, source.ListId) {
            OverrideCount = source.OverrideCount
        };
        foreach (RtfListLevelOverride levelOverride in source.LevelOverrides) {
            listOverride.AddParsedLevelOverride(new RtfListLevelOverride {
                LevelIndex = levelOverride.LevelIndex,
                OverrideFormat = levelOverride.OverrideFormat,
                OverrideStartAt = levelOverride.OverrideStartAt,
                StartAt = levelOverride.StartAt
            });
        }

        return listOverride;
    }

    private static void EnsureListLevel(RtfListDefinition definition, RtfParagraph paragraph) {
        int levelIndex = Math.Min(8, Math.Max(0, paragraph.ListLevel ?? 0));
        if (definition.Levels.Any(level => level.LevelIndex == levelIndex)) {
            return;
        }

        while (definition.Levels.Count < levelIndex) {
            definition.AddLevel(RtfListKind.Decimal);
        }

        RtfListLevel level = definition.AddLevel(paragraph.ListKind);
        level.LeftIndentTwips = paragraph.LeftIndentTwips ?? 720 * (levelIndex + 1);
        level.FirstLineIndentTwips = paragraph.FirstLineIndentTwips ?? -360;
        level.Text = paragraph.ListKind == RtfListKind.Bullet ? "\u2022" : "%" + (levelIndex + 1).ToString(CultureInfo.InvariantCulture) + ".";
        level.Numbers = paragraph.ListKind == RtfListKind.Bullet ? string.Empty : "\u0001";
    }

    private static int ToRtfListLevelAlignmentValue(RtfListLevelAlignment? alignment) {
        switch (alignment) {
            case RtfListLevelAlignment.Center:
                return 1;
            case RtfListLevelAlignment.Right:
                return 2;
            default:
                return 0;
        }
    }

    private static int ToRtfListLevelFollowValue(RtfListLevelFollowCharacter? followCharacter) {
        switch (followCharacter) {
            case RtfListLevelFollowCharacter.Space:
                return 1;
            case RtfListLevelFollowCharacter.Nothing:
                return 2;
            default:
                return 0;
        }
    }

    private static IEnumerable<RtfParagraph> EnumerateParagraphs(RtfDocument document) {
        foreach (RtfParagraph paragraph in EnumerateParagraphs(document.Blocks)) {
            yield return paragraph;
        }

        foreach (RtfHeaderFooter headerFooter in document.HeaderFooters) {
            foreach (RtfParagraph paragraph in headerFooter.Paragraphs) {
                yield return paragraph;
            }
        }
    }

    private static IEnumerable<RtfParagraph> EnumerateParagraphs(IEnumerable<IRtfBlock> blocks) {
        foreach (IRtfBlock block in blocks) {
            if (block is RtfParagraph paragraph) {
                yield return paragraph;
            } else if (block is RtfTable table) {
                foreach (RtfTableRow row in table.Rows) {
                    foreach (RtfTableCell cell in row.Cells) {
                        foreach (RtfParagraph cellParagraph in EnumerateParagraphs(cell.Blocks)) {
                            yield return cellParagraph;
                        }
                    }
                }
            }
        }
    }

    internal sealed class EffectiveListTables {
        public EffectiveListTables(IReadOnlyList<RtfListDefinition> definitions, IReadOnlyList<RtfListOverride> overrides) {
            Definitions = definitions;
            Overrides = overrides;
        }

        public IReadOnlyList<RtfListDefinition> Definitions { get; }

        public IReadOnlyList<RtfListOverride> Overrides { get; }
    }
}
