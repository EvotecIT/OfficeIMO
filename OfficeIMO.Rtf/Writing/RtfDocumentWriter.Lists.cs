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
            var levels = listOverride.LevelOverrides.Select((item, index) => new { Item = item, Index = item.LevelIndex ?? index }).ToArray();
            if (levels.Any(item => item.Index < 0 || item.Index > 8)) throw new InvalidDataException("List override levels must be between 0 and 8.");
            int count = levels.Length == 0 ? 0 : levels.Max(item => item.Index) == 0 ? 1 : 9;
            builder.Append(count.ToString(CultureInfo.InvariantCulture));
            for (int index = 0; index < count; index++) {
                WriteListLevelOverride(builder, levels.LastOrDefault(item => item.Index == index)?.Item ?? new RtfListLevelOverride(), unicodeSkipCount);
            }

            builder.Append(@"\ls");
            builder.Append(listOverride.Id.ToString(CultureInfo.InvariantCulture));
            builder.Append('}');
        }

        builder.Append('}');
    }

    private static void WriteListLevelOverride(StringBuilder builder, RtfListLevelOverride levelOverride, int unicodeSkipCount) {
        builder.Append(@"{\lfolevel");
        AppendOptionalBinary(builder, @"\listoverrideformat", levelOverride.OverrideFormat);
        AppendOptionalBinary(builder, @"\listoverridestartat", levelOverride.OverrideStartAt);
        if (levelOverride.OverrideFormat == true && levelOverride.Formatting != null) {
            RtfListLevel formatting = new RtfCloneContext().Clone(levelOverride.Formatting)!;
            if (levelOverride.OverrideStartAt == true && levelOverride.StartAt.HasValue) formatting.StartAt = levelOverride.StartAt;
            WriteListLevel(builder, formatting, unicodeSkipCount);
        } else {
            AppendOptionalTwips(builder, @"\levelstartat", levelOverride.StartAt);
            if (levelOverride.Formatting != null) WriteListLevel(builder, levelOverride.Formatting, unicodeSkipCount);
        }
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

        string levelText = RtfListTextCodec.EncodeText(level.Text ?? (level.Kind == RtfListKind.Bullet ? "\u2022" : "%" + (level.LevelIndex + 1).ToString(CultureInfo.InvariantCulture) + "."), out string numberOffsets);
        builder.Append(@"{\leveltext");
        WriteListText(builder, levelText, unicodeSkipCount, includeLength: true);
        builder.Append(";}");
        builder.Append(@"{\levelnumbers");
        WriteListText(builder, level.Numbers ?? numberOffsets, unicodeSkipCount, includeLength: false);
        builder.Append(";}");
        AppendOptionalTwips(builder, @"\fi", level.FirstLineIndentTwips);
        AppendOptionalTwips(builder, @"\li", level.LeftIndentTwips);
        builder.Append('}');
    }

    private static void WriteListText(StringBuilder builder, string text, int unicodeSkipCount, bool includeLength) {
        if (includeLength) {
            if (text.Length > 255) throw new InvalidDataException("List marker templates cannot exceed 255 characters.");
            builder.Append(@"\'").Append(text.Length.ToString("x2", CultureInfo.InvariantCulture));
        }
        foreach (char character in text) {
            if (character < 32) builder.Append(@"\'").Append(((int)character).ToString("x2", CultureInfo.InvariantCulture));
            else builder.Append(RtfTextEncoding.EncodeText(character.ToString(), unicodeSkipCount, useNamedCharacters: false));
        }
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
                StartAt = levelOverride.StartAt,
                Formatting = new RtfCloneContext().Clone(levelOverride.Formatting)
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
        level.Numbers = null;
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

    internal static IEnumerable<RtfParagraph> EnumerateParagraphs(RtfDocument document) {
        var pending = new Stack<object>();
        var visited = new HashSet<object>();
        PushListContent(pending, document.Notes);
        PushListContent(pending, document.HeaderFooters);
        PushListContent(pending, document.Blocks);
        while (pending.Count > 0) {
            object current = pending.Pop();
            if (!visited.Add(current)) continue;
            switch (current) {
                case RtfParagraph paragraph:
                    yield return paragraph;
                    PushListContent(pending, paragraph.Inlines);
                    if (paragraph.ListText != null) pending.Push(paragraph.ListText);
                    break;
                case RtfTable table:
                    PushListContent(pending, table.Rows);
                    break;
                case RtfTableRow row:
                    PushListContent(pending, row.Cells);
                    break;
                case RtfTableCell cell:
                    PushListContent(pending, cell.Blocks);
                    break;
                case RtfHeaderFooter headerFooter:
                    PushListContent(pending, headerFooter.Paragraphs);
                    break;
                case RtfNote note:
                    PushListContent(pending, note.Paragraphs);
                    break;
                case RtfRun run when run.Note != null:
                    pending.Push(run.Note);
                    break;
                case RtfGeneratedText generated when generated.Note != null:
                    pending.Push(generated.Note);
                    break;
                case RtfField field:
                    pending.Push(field.Result);
                    break;
                case RtfObject rtfObject:
                    pending.Push(rtfObject.Result);
                    break;
                case RtfShape shape:
                    PushListContent(pending, shape.TextBoxParagraphs);
                    break;
            }
        }
    }

    private static void PushListContent<T>(Stack<object> pending, IReadOnlyList<T> content) where T : class {
        for (int index = content.Count - 1; index >= 0; index--) pending.Push(content[index]);
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
