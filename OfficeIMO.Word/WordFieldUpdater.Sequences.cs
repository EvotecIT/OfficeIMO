using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal static partial class WordFieldUpdater {
        private static bool TryEvaluateSequence(
            WordDocument document,
            MutableFieldCandidate candidate,
            WordFieldInventory.ParsedFieldInstruction parsed,
            FieldEvaluationState state,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Skipped;

            string? sequenceName = parsed.Instructions.FirstOrDefault();
            if (string.IsNullOrWhiteSpace(sequenceName)) {
                message = "SEQ field is missing a sequence identifier.";
                return false;
            }

            sequenceName = TrimQuotes(sequenceName);
            if (!TryGetSequenceSwitches(parsed, out SequenceSwitches switches, out string? unsupportedSwitch, out string? switchError)) {
                status = WordFieldUpdateStatus.Unsupported;
                message = switchError ?? $"SEQ field switch {unsupportedSwitch} is not evaluated by OfficeIMO.";
                return false;
            }

            if (switches.HeadingResetLevel.HasValue && !TryApplyHeadingReset(document, candidate, state, sequenceName, switches.HeadingResetLevel.Value)) {
                status = WordFieldUpdateStatus.Skipped;
                message = $"SEQ field sequence {sequenceName} could not find a preceding Heading {switches.HeadingResetLevel.Value} paragraph for the \\s reset switch.";
                return false;
            }

            int currentValue;
            if (switches.ResetValue.HasValue) {
                currentValue = switches.ResetValue.Value;
                state.Sequences[sequenceName] = currentValue;
            } else if (switches.RepeatCurrent) {
                if (!state.Sequences.TryGetValue(sequenceName, out currentValue)) {
                    message = $"SEQ field sequence {sequenceName} has no previous value to repeat.";
                    return false;
                }
            } else {
                state.Sequences.TryGetValue(sequenceName, out currentValue);
                currentValue++;
                state.Sequences[sequenceName] = currentValue;
            }

            value = FormatSequenceValue(currentValue, parsed.FormatSwitches);
            status = WordFieldUpdateStatus.Updated;
            message = switches.ResetValue.HasValue
                ? $"Updated sequence {sequenceName} from reset value {currentValue.ToString(CultureInfo.InvariantCulture)}."
                : switches.RepeatCurrent
                    ? $"Updated sequence {sequenceName} from the current sequence value."
                    : switches.HeadingResetLevel.HasValue
                        ? $"Updated sequence {sequenceName} from Heading {switches.HeadingResetLevel.Value} reset context."
                    : $"Updated sequence {sequenceName} from document order.";
            return true;
        }

        private static bool TryGetSequenceSwitches(
            WordFieldInventory.ParsedFieldInstruction parsed,
            out SequenceSwitches switches,
            out string? unsupportedSwitch,
            out string? error) {
            int? resetValue = null;
            bool repeatCurrent = false;
            int? headingResetLevel = null;
            unsupportedSwitch = null;
            error = null;

            foreach (string fieldSwitch in parsed.Switches) {
                string trimmed = fieldSwitch.Trim();
                if (string.Equals(trimmed, "\\c", StringComparison.OrdinalIgnoreCase)) {
                    repeatCurrent = true;
                    continue;
                }

                if (string.Equals(trimmed, "\\n", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(trimmed, "\\h", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                if (trimmed.StartsWith("\\s", StringComparison.OrdinalIgnoreCase)) {
                    string rawValue = trimmed.Substring(2).Trim();
                    rawValue = TrimQuotes(rawValue);
                    if (!int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedValue) ||
                        parsedValue < 1 ||
                        parsedValue > 9) {
                        error = $"SEQ field heading reset switch {trimmed} must use a heading level from 1 to 9.";
                        switches = default;
                        return false;
                    }

                    headingResetLevel = parsedValue;
                    continue;
                }

                if (trimmed.StartsWith("\\r", StringComparison.OrdinalIgnoreCase)) {
                    string rawValue = trimmed.Substring(2).Trim();
                    rawValue = TrimQuotes(rawValue);
                    if (!int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedValue) || parsedValue < 0) {
                        error = $"SEQ field reset switch {trimmed} must use a non-negative integer.";
                        switches = default;
                        return false;
                    }

                    resetValue = parsedValue;
                    continue;
                }

                unsupportedSwitch = trimmed;
                switches = default;
                return false;
            }

            if (repeatCurrent && resetValue.HasValue) {
                error = "SEQ field cannot combine the repeat-current switch with a reset switch.";
                switches = default;
                return false;
            }

            if (resetValue.HasValue && headingResetLevel.HasValue) {
                error = "SEQ field cannot combine an explicit reset switch with a heading reset switch.";
                switches = default;
                return false;
            }

            if (repeatCurrent && headingResetLevel.HasValue) {
                error = "SEQ field cannot combine the repeat-current switch with a heading reset switch.";
                switches = default;
                return false;
            }

            switches = new SequenceSwitches(resetValue, repeatCurrent, headingResetLevel);
            return true;
        }

        private static bool TryApplyHeadingReset(
            WordDocument document,
            MutableFieldCandidate candidate,
            FieldEvaluationState state,
            string sequenceName,
            int headingLevel) {
            string? headingKey = FindNearestPrecedingHeadingKey(document, candidate.AnchorElement, headingLevel);
            if (headingKey == null) {
                return false;
            }

            string resetKey = sequenceName + "|" + headingLevel.ToString(CultureInfo.InvariantCulture);
            if (!state.SequenceHeadingResetKeys.TryGetValue(resetKey, out string? knownHeadingKey) ||
                !string.Equals(knownHeadingKey, headingKey, StringComparison.Ordinal)) {
                state.SequenceHeadingResetKeys[resetKey] = headingKey;
                state.Sequences[sequenceName] = 0;
            }

            return true;
        }

        private static string? FindNearestPrecedingHeadingKey(WordDocument document, OpenXmlElement anchorElement, int headingLevel) {
            Body? body = document._wordprocessingDocument.MainDocumentPart?.Document?.Body;
            Paragraph? targetParagraph = anchorElement is Paragraph paragraph
                ? paragraph
                : anchorElement.Ancestors<Paragraph>().FirstOrDefault();

            if (body == null || targetParagraph == null) {
                return null;
            }

            string? currentHeadingKey = null;
            int paragraphIndex = 0;
            foreach (Paragraph currentParagraph in body.Descendants<Paragraph>()) {
                if (ReferenceEquals(currentParagraph, targetParagraph)) {
                    return currentHeadingKey;
                }

                if (TryGetHeadingLevel(currentParagraph, out int currentLevel) && currentLevel == headingLevel) {
                    currentHeadingKey = paragraphIndex.ToString(CultureInfo.InvariantCulture);
                }

                paragraphIndex++;
            }

            return null;
        }

        private static bool TryGetHeadingLevel(Paragraph paragraph, out int level) {
            level = 0;
            string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (string.IsNullOrWhiteSpace(styleId)) {
                return false;
            }

            string normalized = styleId!.Replace(" ", string.Empty);
            if (!normalized.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            return int.TryParse(normalized.Substring("Heading".Length), NumberStyles.Integer, CultureInfo.InvariantCulture, out level) &&
                level >= 1 &&
                level <= 9;
        }

        private static string FormatSequenceValue(int value, IReadOnlyList<WordFieldFormat> formatSwitches) {
            WordFieldFormat? format = formatSwitches
                .Where(item => item != WordFieldFormat.Mergeformat)
                .Cast<WordFieldFormat?>()
                .LastOrDefault();

            switch (format) {
                case WordFieldFormat.Roman:
                    return ToRoman(value).ToUpperInvariant();
                case WordFieldFormat.roman:
                    return ToRoman(value).ToLowerInvariant();
                case WordFieldFormat.Ordinal:
                    return ToOrdinal(value);
                case WordFieldFormat.Alphabetical:
                    return ToAlphabetic(value, uppercase: false);
                case WordFieldFormat.ALPHABETICAL:
                    return ToAlphabetic(value, uppercase: true);
                case WordFieldFormat.Hex:
                    return ToHex(value);
                case WordFieldFormat.CardText:
                    return ToCardinalText(value);
                case WordFieldFormat.OrdText:
                    return ToOrdinalText(value);
                case WordFieldFormat.DollarText:
                    return ToDollarText(value);
                default:
                    return value.ToString(CultureInfo.InvariantCulture);
            }
        }

        private static string ToRoman(int value) {
            if (value <= 0 || value > 3999) {
                return value.ToString(CultureInfo.InvariantCulture);
            }

            (int Number, string Numeral)[] map = {
                (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
                (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
                (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
            };

            var builder = new System.Text.StringBuilder();
            int remaining = value;
            foreach ((int number, string numeral) in map) {
                while (remaining >= number) {
                    builder.Append(numeral);
                    remaining -= number;
                }
            }

            return builder.ToString();
        }

        private static string ToOrdinal(int value) {
            int absoluteValue = Math.Abs(value);
            int lastTwoDigits = absoluteValue % 100;
            string suffix;
            if (lastTwoDigits is >= 11 and <= 13) {
                suffix = "th";
            } else {
                suffix = (absoluteValue % 10) switch {
                    1 => "st",
                    2 => "nd",
                    3 => "rd",
                    _ => "th"
                };
            }

            return value.ToString(CultureInfo.InvariantCulture) + suffix;
        }

        private static string ToHex(int value) {
            return value < 0
                ? value.ToString(CultureInfo.InvariantCulture)
                : value.ToString("X", CultureInfo.InvariantCulture);
        }

        private static string ToDollarText(int value) {
            if (value < 0) {
                return value.ToString(CultureInfo.InvariantCulture);
            }

            return ToCardinalText(value) + " and 00/100";
        }

        private static string ToOrdinalText(int value) {
            if (value < 0) {
                return value.ToString(CultureInfo.InvariantCulture);
            }

            return ToOrdinalText((long)value);
        }

        private static string ToOrdinalText(long value) {
            if (value < 20) {
                return value switch {
                    0 => "zeroth",
                    1 => "first",
                    2 => "second",
                    3 => "third",
                    4 => "fourth",
                    5 => "fifth",
                    6 => "sixth",
                    7 => "seventh",
                    8 => "eighth",
                    9 => "ninth",
                    10 => "tenth",
                    11 => "eleventh",
                    12 => "twelfth",
                    13 => "thirteenth",
                    14 => "fourteenth",
                    15 => "fifteenth",
                    16 => "sixteenth",
                    17 => "seventeenth",
                    18 => "eighteenth",
                    _ => "nineteenth"
                };
            }

            if (value < 100) {
                long tens = value / 10;
                long remainder = value % 10;
                string tensText = tens switch {
                    2 => "twent",
                    3 => "thirt",
                    4 => "fort",
                    5 => "fift",
                    6 => "sixt",
                    7 => "sevent",
                    8 => "eight",
                    _ => "ninet"
                };

                return remainder == 0
                    ? tensText + "ieth"
                    : ToCardinalText(tens * 10) + "-" + ToOrdinalText(remainder);
            }

            if (value < 1000) {
                long hundreds = value / 100;
                long remainder = value % 100;
                return remainder == 0
                    ? ToCardinalText(hundreds) + " hundredth"
                    : ToCardinalText(hundreds) + " hundred " + ToOrdinalText(remainder);
            }

            return ToScaledOrdinalText(value);
        }

        private static string ToScaledOrdinalText(long value) {
            (long Scale, string Name, string OrdinalName)[] scales = {
                (1000000000, "billion", "billionth"),
                (1000000, "million", "millionth"),
                (1000, "thousand", "thousandth")
            };

            foreach ((long scale, string name, string ordinalName) in scales) {
                if (value < scale) {
                    continue;
                }

                long leading = value / scale;
                long remainder = value % scale;
                return remainder == 0
                    ? ToCardinalText(leading) + " " + ordinalName
                    : ToCardinalText(leading) + " " + name + " " + ToOrdinalText(remainder);
            }

            return value.ToString(CultureInfo.InvariantCulture);
        }

        private static string ToCardinalText(long value) {
            if (value < 0) {
                return value.ToString(CultureInfo.InvariantCulture);
            }

            if (value < 20) {
                return value switch {
                    0 => "zero",
                    1 => "one",
                    2 => "two",
                    3 => "three",
                    4 => "four",
                    5 => "five",
                    6 => "six",
                    7 => "seven",
                    8 => "eight",
                    9 => "nine",
                    10 => "ten",
                    11 => "eleven",
                    12 => "twelve",
                    13 => "thirteen",
                    14 => "fourteen",
                    15 => "fifteen",
                    16 => "sixteen",
                    17 => "seventeen",
                    18 => "eighteen",
                    _ => "nineteen"
                };
            }

            if (value < 100) {
                long tens = value / 10;
                long remainder = value % 10;
                string text = tens switch {
                    2 => "twenty",
                    3 => "thirty",
                    4 => "forty",
                    5 => "fifty",
                    6 => "sixty",
                    7 => "seventy",
                    8 => "eighty",
                    _ => "ninety"
                };

                return remainder == 0 ? text : text + "-" + ToCardinalText(remainder);
            }

            if (value < 1000) {
                long hundreds = value / 100;
                long remainder = value % 100;
                string text = ToCardinalText(hundreds) + " hundred";
                return remainder == 0 ? text : text + " " + ToCardinalText(remainder);
            }

            return ToScaledCardinalText(value);
        }

        private static string ToScaledCardinalText(long value) {
            (long Scale, string Name)[] scales = {
                (1000000000, "billion"),
                (1000000, "million"),
                (1000, "thousand")
            };

            foreach ((long scale, string name) in scales) {
                if (value < scale) {
                    continue;
                }

                long leading = value / scale;
                long remainder = value % scale;
                string text = ToCardinalText(leading) + " " + name;
                return remainder == 0 ? text : text + " " + ToCardinalText(remainder);
            }

            return value.ToString(CultureInfo.InvariantCulture);
        }

        private static string ToAlphabetic(int value, bool uppercase) {
            if (value <= 0) {
                return value.ToString(CultureInfo.InvariantCulture);
            }

            var chars = new Stack<char>();
            int remaining = value;
            while (remaining > 0) {
                remaining--;
                chars.Push((char)((uppercase ? 'A' : 'a') + remaining % 26));
                remaining /= 26;
            }

            return new string(chars.ToArray());
        }

        private readonly struct SequenceSwitches {
            internal SequenceSwitches(int? resetValue, bool repeatCurrent, int? headingResetLevel) {
                ResetValue = resetValue;
                RepeatCurrent = repeatCurrent;
                HeadingResetLevel = headingResetLevel;
            }

            internal int? ResetValue { get; }

            internal bool RepeatCurrent { get; }

            internal int? HeadingResetLevel { get; }
        }
    }
}
