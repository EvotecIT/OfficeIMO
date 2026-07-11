using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static Dictionary<int, RtfListDefinition> CreateListDefinitionLookup(IReadOnlyList<RtfListDefinition> definitions) {
            return definitions.GroupBy(definition => definition.Id).ToDictionary(group => group.Key, group => group.First());
        }

        private static Dictionary<int, RtfListOverride> CreateListOverrideLookup(IReadOnlyList<RtfListOverride> overrides) {
            return overrides.GroupBy(listOverride => listOverride.Id).ToDictionary(group => group.Key, group => group.First());
        }

        private void ApplyListOverride(CharacterState state) {
            state.ListDefinitionId = null;
            if (!state.ListId.HasValue || !_listOverridesById.TryGetValue(state.ListId.Value, out RtfListOverride? listOverride)) {
                return;
            }

            state.ListDefinitionId = listOverride.ListId;
            ApplyListLevel(state);
        }

        private void ApplyListLevel(CharacterState state) {
            if (!state.ListDefinitionId.HasValue ||
                !_listDefinitionsById.TryGetValue(state.ListDefinitionId.Value, out RtfListDefinition? definition)) {
                return;
            }

            int levelIndex = state.ListLevel ?? 0;
            RtfListLevel? level = definition.Levels.FirstOrDefault(item => item.LevelIndex == levelIndex) ??
                                  definition.Levels.ElementAtOrDefault(levelIndex);
            if (level == null) {
                return;
            }

            state.ListKind = level.Kind;
            state.LeftIndentTwips ??= level.LeftIndentTwips;
            state.FirstLineIndentTwips ??= level.FirstLineIndentTwips;
        }

        private static IReadOnlyList<RtfListDefinition> ReadListDefinitions(RtfGroup root, int ansiCodePage, int unicodeSkipCount) {
            RtfGroup? listTable = root.Children.OfType<RtfGroup>().FirstOrDefault(group => group.Destination == "listtable");
            if (listTable == null) return Array.Empty<RtfListDefinition>();

            var definitions = new List<RtfListDefinition>();
            foreach (RtfGroup listGroup in listTable.Children.OfType<RtfGroup>().Where(group => group.Destination == "list")) {
                RtfListDefinition? definition = ReadListDefinition(listGroup, ansiCodePage, unicodeSkipCount);
                if (definition != null) {
                    definitions.Add(definition);
                }
            }

            return definitions;
        }

        private static RtfListDefinition? ReadListDefinition(RtfGroup listGroup, int ansiCodePage, int unicodeSkipCount) {
            int? listId = null;
            int? templateId = null;
            string? name = null;
            var levels = new List<RtfListLevel>();

            foreach (RtfNode node in listGroup.Children) {
                if (node is RtfControlWord control) {
                    switch (control.Name) {
                        case "listid":
                            listId = control.Parameter;
                            break;
                        case "listtemplateid":
                            templateId = control.Parameter;
                            break;
                    }
                } else if (node is RtfGroup group) {
                    if (group.Destination == "listlevel") {
                        levels.Add(ReadListLevel(group, levels.Count, ansiCodePage, unicodeSkipCount));
                    } else if (group.Destination == "listname") {
                        name = CleanListText(CollectPlainText(group, ansiCodePage, unicodeSkipCount));
                    }
                }
            }

            if (!listId.HasValue) return null;

            var definition = new RtfListDefinition(listId.Value) {
                TemplateId = templateId,
                Name = name
            };
            foreach (RtfListLevel level in levels) {
                definition.AddParsedLevel(level);
            }

            return definition;
        }

        private static RtfListLevel ReadListLevel(RtfGroup levelGroup, int levelIndex, int ansiCodePage, int unicodeSkipCount) {
            int? numberFormat = null;
            int? numberFormatN = null;
            int? startAt = null;
            RtfListLevelAlignment? alignment = null;
            RtfListLevelAlignment? alignmentN = null;
            RtfListLevelFollowCharacter? followCharacter = null;
            int? space = null;
            int? indent = null;
            bool? legalNumbering = null;
            bool? noRestart = null;
            int? pictureIndex = null;
            bool pictureNoSize = false;
            int? leftIndent = null;
            int? firstLineIndent = null;
            string? levelText = null;
            string? levelNumbers = null;

            foreach (RtfNode node in levelGroup.Children) {
                if (node is RtfControlWord control) {
                    switch (control.Name) {
                        case "levelnfc":
                            numberFormat = control.Parameter;
                            break;
                        case "levelnfcn":
                            numberFormatN = control.Parameter;
                            break;
                        case "leveljc":
                            alignment = ToListLevelAlignment(control.Parameter);
                            break;
                        case "leveljcn":
                            alignmentN = ToListLevelAlignment(control.Parameter);
                            break;
                        case "levelfollow":
                            followCharacter = ToListLevelFollowCharacter(control.Parameter);
                            break;
                        case "levelstartat":
                            startAt = control.Parameter;
                            break;
                        case "levelspace":
                            space = control.Parameter;
                            break;
                        case "levelindent":
                            indent = control.Parameter;
                            break;
                        case "levellegal":
                            legalNumbering = !control.HasParameter || control.Parameter != 0;
                            break;
                        case "levelnorestart":
                            noRestart = !control.HasParameter || control.Parameter != 0;
                            break;
                        case "levelpicture":
                            pictureIndex = control.Parameter;
                            break;
                        case "levelpicturenosize":
                            pictureNoSize = !control.HasParameter || control.Parameter != 0;
                            break;
                        case "li":
                            leftIndent = control.Parameter;
                            break;
                        case "fi":
                            firstLineIndent = control.Parameter;
                            break;
                    }
                } else if (node is RtfGroup group) {
                    if (group.Destination == "leveltext") {
                        levelText = CleanListText(CollectPlainText(group, ansiCodePage, unicodeSkipCount));
                    } else if (group.Destination == "levelnumbers") {
                        levelNumbers = CleanListText(CollectPlainText(group, ansiCodePage, unicodeSkipCount));
                    }
                }
            }

            var level = new RtfListLevel(levelIndex, GetListKind(numberFormatN ?? numberFormat, levelText)) {
                NumberFormat = numberFormat,
                NumberFormatN = numberFormatN,
                Alignment = alignment,
                AlignmentN = alignmentN,
                FollowCharacter = followCharacter,
                StartAt = startAt,
                SpaceTwips = space,
                IndentTwips = indent,
                LegalNumbering = legalNumbering,
                NoRestart = noRestart,
                PictureIndex = pictureIndex,
                PictureNoSize = pictureNoSize,
                Text = levelText,
                Numbers = levelNumbers,
                LeftIndentTwips = leftIndent,
                FirstLineIndentTwips = firstLineIndent
            };
            return level;
        }

        private static RtfListLevelAlignment? ToListLevelAlignment(int? value) {
            switch (value) {
                case 0:
                    return RtfListLevelAlignment.Left;
                case 1:
                    return RtfListLevelAlignment.Center;
                case 2:
                    return RtfListLevelAlignment.Right;
                default:
                    return null;
            }
        }

        private static RtfListLevelFollowCharacter? ToListLevelFollowCharacter(int? value) {
            switch (value) {
                case 0:
                    return RtfListLevelFollowCharacter.Tab;
                case 1:
                    return RtfListLevelFollowCharacter.Space;
                case 2:
                    return RtfListLevelFollowCharacter.Nothing;
                default:
                    return null;
            }
        }

        private static IReadOnlyList<RtfListOverride> ReadListOverrides(RtfGroup root) {
            RtfGroup? overrideTable = root.Children.OfType<RtfGroup>().FirstOrDefault(group => group.Destination == "listoverridetable");
            if (overrideTable == null) return Array.Empty<RtfListOverride>();

            var overrides = new List<RtfListOverride>();
            foreach (RtfGroup overrideGroup in overrideTable.Children.OfType<RtfGroup>().Where(group => group.Destination == "listoverride")) {
                int? listId = null;
                int? overrideId = null;
                int? overrideCount = null;
                var levelOverrides = new List<RtfListLevelOverride>();
                foreach (RtfNode node in overrideGroup.Children) {
                    if (node is RtfControlWord control) {
                        switch (control.Name) {
                            case "listid":
                                listId = control.Parameter;
                                break;
                            case "ls":
                                overrideId = control.Parameter;
                                break;
                            case "listoverridecount":
                                overrideCount = control.Parameter;
                                break;
                        }
                    } else if (node is RtfGroup group && group.Destination == "lfolevel") {
                        RtfListLevelOverride? levelOverride = ReadListLevelOverride(group);
                        if (levelOverride != null) {
                            levelOverrides.Add(levelOverride);
                        }
                    }
                }

                if (overrideId.HasValue && listId.HasValue) {
                    var listOverride = new RtfListOverride(overrideId.Value, listId.Value) {
                        OverrideCount = overrideCount
                    };
                    foreach (RtfListLevelOverride levelOverride in levelOverrides) {
                        levelOverride.LevelIndex = listOverride.LevelOverrides.Count;
                        listOverride.AddParsedLevelOverride(levelOverride);
                    }

                    overrides.Add(listOverride);
                }
            }

            return overrides;
        }

        private static RtfListLevelOverride? ReadListLevelOverride(RtfGroup group) {
            var levelOverride = new RtfListLevelOverride();
            foreach (RtfNode node in group.Children) {
                if (!(node is RtfControlWord control)) {
                    continue;
                }

                switch (control.Name) {
                    case "listoverrideformat":
                        levelOverride.OverrideFormat = !control.HasParameter || control.Parameter != 0;
                        break;
                    case "listoverridestartat":
                        levelOverride.OverrideStartAt = !control.HasParameter || control.Parameter != 0;
                        break;
                    case "levelstartat":
                        levelOverride.StartAt = control.Parameter;
                        break;
                }
            }

            return levelOverride.HasAnyValue ? levelOverride : null;
        }

        private static RtfListKind GetListKind(int? numberFormat, string? levelText) {
            if (numberFormat == 23 || (levelText != null && levelText.IndexOf('\u2022') >= 0)) {
                return RtfListKind.Bullet;
            }

            return RtfListKind.Decimal;
        }

        private static string CleanListText(string text) {
            string cleaned = text.Trim().TrimEnd(';').Trim();
            while (cleaned.Length > 0 && char.IsControl(cleaned[0])) {
                cleaned = cleaned.Substring(1);
            }

            return cleaned.Trim();
        }
    }
}
