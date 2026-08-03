using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal static partial class WordFieldUpdater {
        private static bool TryEvaluateListNum(
            WordDocument document,
            MutableFieldCandidate candidate,
            WordFieldInventory.ParsedFieldInstruction parsed,
            FieldEvaluationState state,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Unsupported;

            if (candidate.NestingLevel > 0) {
                message = "Nested LISTNUM fields are ignored by Word and were left unchanged.";
                return false;
            }

            if (!TryGetListNumProfile(parsed, out ListNumProfile profile, out string profileName, out string? profileError)) {
                message = profileError!;
                return false;
            }

            if (!TryGetListNumSwitches(parsed, out int? explicitLevel, out int? startAt, out string? switchError)) {
                message = switchError!;
                return false;
            }

            int level = explicitLevel ?? GetListNumParagraphLevel(document, candidate.AnchorElement);
            string stateKey = candidate.PartUri + "|" + profileName;
            if (!state.ListNumSequences.TryGetValue(stateKey, out ListNumSequenceState? sequence)) {
                sequence = new ListNumSequenceState();
                state.ListNumSequences[stateKey] = sequence;
            }

            int levelIndex = level - 1;
            if (startAt.HasValue) {
                sequence.Counters[levelIndex] = startAt.Value;
            } else {
                if (sequence.Counters[levelIndex] == int.MaxValue) {
                    message = $"{profileName} LISTNUM level {level.ToString(CultureInfo.InvariantCulture)} cannot advance beyond the supported integer range.";
                    return false;
                }

                sequence.Counters[levelIndex] = sequence.Counters[levelIndex] <= 0
                    ? 1
                    : sequence.Counters[levelIndex] + 1;
            }

            for (int index = levelIndex + 1; index < sequence.Counters.Length; index++) {
                sequence.Counters[index] = 0;
            }

            if (profile == ListNumProfile.LegalDefault) {
                for (int index = 0; index < levelIndex; index++) {
                    if (sequence.Counters[index] <= 0) {
                        sequence.Counters[index] = 1;
                    }
                }
            }

            value = FormatListNum(profile, levelIndex, sequence.Counters);
            status = WordFieldUpdateStatus.Updated;
            message = startAt.HasValue
                ? $"Updated {profileName} LISTNUM level {level.ToString(CultureInfo.InvariantCulture)} from start value {startAt.Value.ToString(CultureInfo.InvariantCulture)}."
                : $"Updated {profileName} LISTNUM level {level.ToString(CultureInfo.InvariantCulture)} from document order.";
            return true;
        }

        private static bool TryGetListNumProfile(
            WordFieldInventory.ParsedFieldInstruction parsed,
            out ListNumProfile profile,
            out string profileName,
            out string? error) {
            profileName = parsed.Instructions.Count == 0
                ? "NumberDefault"
                : TrimQuotes(parsed.Instructions[0]);

            if (parsed.Instructions.Count > 1) {
                profile = default;
                error = "LISTNUM accepts at most one list-template name.";
                return false;
            }

            if (string.Equals(profileName, "NumberDefault", StringComparison.OrdinalIgnoreCase)) {
                profile = ListNumProfile.NumberDefault;
                profileName = "NumberDefault";
                error = null;
                return true;
            }

            if (string.Equals(profileName, "OutlineDefault", StringComparison.OrdinalIgnoreCase)) {
                profile = ListNumProfile.OutlineDefault;
                profileName = "OutlineDefault";
                error = null;
                return true;
            }

            if (string.Equals(profileName, "LegalDefault", StringComparison.OrdinalIgnoreCase)) {
                profile = ListNumProfile.LegalDefault;
                profileName = "LegalDefault";
                error = null;
                return true;
            }

            profile = default;
            error = $"LISTNUM list template {profileName} is not one of the deterministic built-in profiles NumberDefault, OutlineDefault, or LegalDefault.";
            return false;
        }

        private static bool TryGetListNumSwitches(
            WordFieldInventory.ParsedFieldInstruction parsed,
            out int? level,
            out int? startAt,
            out string? error) {
            level = null;
            startAt = null;
            error = null;

            if (!string.IsNullOrWhiteSpace(parsed.NumericPictureSwitch)) {
                error = "LISTNUM numeric picture switches are not part of the deterministic built-in profiles.";
                return false;
            }

            if (parsed.FormatSwitches.Any(format => format is not WordFieldFormat.Mergeformat and not WordFieldFormat.CharFormat)) {
                error = "LISTNUM supports only Mergeformat and CharFormat general format switches in deterministic built-in profiles.";
                return false;
            }

            foreach (string fieldSwitch in parsed.Switches) {
                string trimmed = fieldSwitch.Trim();
                if (trimmed.StartsWith("\\l", StringComparison.OrdinalIgnoreCase)) {
                    if (level.HasValue || !TryParseListNumSwitchValue(trimmed, out int parsedLevel) || parsedLevel < 1 || parsedLevel > 9) {
                        error = $"LISTNUM level switch {trimmed} must occur once and use a level from 1 to 9.";
                        return false;
                    }

                    level = parsedLevel;
                    continue;
                }

                if (trimmed.StartsWith("\\s", StringComparison.OrdinalIgnoreCase)) {
                    if (startAt.HasValue || !TryParseListNumSwitchValue(trimmed, out int parsedStart) || parsedStart < 1) {
                        error = $"LISTNUM start-at switch {trimmed} must occur once and use a positive integer.";
                        return false;
                    }

                    startAt = parsedStart;
                    continue;
                }

                error = $"LISTNUM switch {trimmed} is not supported by the deterministic built-in profiles.";
                return false;
            }

            return true;
        }

        private static bool TryParseListNumSwitchValue(string fieldSwitch, out int value) {
            string rawValue = fieldSwitch.Length > 2 ? TrimQuotes(fieldSwitch.Substring(2).Trim()) : string.Empty;
            return int.TryParse(rawValue, NumberStyles.None, CultureInfo.InvariantCulture, out value);
        }

        private static int GetListNumParagraphLevel(WordDocument document, DocumentFormat.OpenXml.OpenXmlElement anchorElement) {
            Paragraph? paragraph = anchorElement is Paragraph directParagraph
                ? directParagraph
                : anchorElement.Ancestors<Paragraph>().FirstOrDefault();
            if (paragraph == null) {
                return 1;
            }

            int? directLevel = paragraph.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value;
            if (directLevel is >= 0 and <= 8) {
                return directLevel.Value + 1;
            }

            Numbering? numbering = document._wordprocessingDocument.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
            if (numbering != null) {
                IReadOnlyDictionary<string, ReferenceParagraphNumbering> styles = BuildReferenceParagraphStyleNumbering(document, numbering);
                if (TryGetParagraphNumbering(paragraph, styles, out _, out int styleLevel) && styleLevel is >= 0 and <= 8) {
                    return styleLevel + 1;
                }
            }

            return 1;
        }

        private static string FormatListNum(ListNumProfile profile, int levelIndex, IReadOnlyList<int> counters) {
            int current = counters[levelIndex];
            if (profile == ListNumProfile.LegalDefault) {
                return string.Join(".", counters.Take(levelIndex + 1).Select(counter => counter.ToString(CultureInfo.InvariantCulture))) + ".";
            }

            return profile == ListNumProfile.NumberDefault
                ? levelIndex switch {
                    0 => current.ToString(CultureInfo.InvariantCulture) + ")",
                    1 => ToAlphabetic(current, uppercase: false) + ")",
                    2 => ToRoman(current).ToLowerInvariant() + ")",
                    3 => "(" + current.ToString(CultureInfo.InvariantCulture) + ")",
                    4 => "(" + ToAlphabetic(current, uppercase: false) + ")",
                    5 => "(" + ToRoman(current).ToLowerInvariant() + ")",
                    6 => current.ToString(CultureInfo.InvariantCulture) + ".",
                    7 => ToAlphabetic(current, uppercase: false) + ".",
                    _ => ToRoman(current).ToLowerInvariant() + "."
                }
                : levelIndex switch {
                    0 => ToRoman(current).ToUpperInvariant() + ".",
                    1 => ToAlphabetic(current, uppercase: true) + ".",
                    2 => current.ToString(CultureInfo.InvariantCulture) + ".",
                    3 => ToAlphabetic(current, uppercase: false) + ")",
                    4 => "(" + current.ToString(CultureInfo.InvariantCulture) + ")",
                    5 => "(" + ToAlphabetic(current, uppercase: false) + ")",
                    6 => "(" + ToRoman(current).ToLowerInvariant() + ")",
                    7 => "(" + ToAlphabetic(current, uppercase: false) + ")",
                    _ => "(" + ToRoman(current).ToLowerInvariant() + ")"
                };
        }

        private enum ListNumProfile {
            NumberDefault,
            OutlineDefault,
            LegalDefault
        }

        private sealed class ListNumSequenceState {
            internal int[] Counters { get; } = new int[9];
        }
    }
}
