using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    public static partial class WordDocumentTraversal {
        private const int MaximumListMarkerLength = 4096;
        private const int MaximumDocumentListMarkerCharacters = 4 * 1024 * 1024;

        private static string BuildMarker(int level, int index, Dictionary<int, int> indices, Dictionary<int, WordNumberFormat?> formats, string? pattern) {
            if (string.IsNullOrEmpty(pattern)) {
                string formatted = FormatNumber(index, formats[level]);
                return formatted + ".";
            }

            if (pattern!.Length > MaximumListMarkerLength) throw ListMarkerLengthExceeded();
            var marker = new StringBuilder(Math.Min(pattern.Length + 16, MaximumListMarkerLength));
            int consumed = 0;
            foreach (Match match in Regex.Matches(pattern, "%CurrentLevel|%([0-9]+)")) {
                Append(pattern.Substring(consumed, match.Index - consumed));
                if (match.Value == "%CurrentLevel") Append(FormatNumber(index, formats[level]));
                else if (int.TryParse(match.Groups[1].Value, out int placeholderLevel) && placeholderLevel > 0) {
                    int lvl = placeholderLevel - 1;
                    int value = lvl == level ? index : indices.TryGetValue(lvl, out int val) ? val - 1 : 0;
                    formats.TryGetValue(lvl, out WordNumberFormat? fmt);
                    Append(FormatNumber(value, fmt));
                } else Append(match.Value);
                consumed = match.Index + match.Length;
            }
            Append(pattern.Substring(consumed));
            return marker.ToString();

            void Append(string value) {
                if (value.Length > MaximumListMarkerLength - marker.Length) throw ListMarkerLengthExceeded();
                marker.Append(value);
            }
        }

        private static InvalidDataException ListMarkerLengthExceeded() =>
            new("Word list marker exceeds the supported length of 4096 characters.");

        private static string FormatNumber(int number, WordNumberFormat? format) {
            if (format == WordNumberFormat.LowerRoman) {
                return ToRoman(number).ToLowerInvariant();
            }
            if (format == WordNumberFormat.UpperRoman) {
                return ToRoman(number);
            }
            if (format == WordNumberFormat.LowerLetter) {
                return ToAlphabeticSequence(number, uppercase: false);
            }
            if (format == WordNumberFormat.UpperLetter) {
                return ToAlphabeticSequence(number, uppercase: true);
            }
            return number.ToString();
        }

        private static string ToAlphabeticSequence(int number, bool uppercase) {
            if (number <= 0) {
                return number.ToString();
            }

            char baseCharacter = uppercase ? 'A' : 'a';
            StringBuilder sb = new();
            while (number > 0) {
                number--;
                sb.Insert(0, (char)(baseCharacter + (number % 26)));
                number /= 26;
            }

            return sb.ToString();
        }

        private static string ToRoman(int number) {
            if (number <= 0) {
                return number.ToString();
            }

            (int Value, string Symbol)[] map = new (int, string)[] {
                (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
                (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
                (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
            };

            int remaining = number;
            int length = 0;
            foreach (var (value, symbol) in map) {
                length += (remaining / value) * symbol.Length;
                remaining %= value;
            }
            if (length > MaximumListMarkerLength) throw ListMarkerLengthExceeded();

            StringBuilder sb = new(length);
            foreach ((int value, string symbol) in map) {
                while (number >= value) {
                    sb.Append(symbol);
                    number -= value;
                }
            }

            return sb.ToString();
        }
    }
}
