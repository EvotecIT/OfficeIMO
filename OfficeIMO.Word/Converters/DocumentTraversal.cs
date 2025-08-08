using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides helper methods for traversing documents and resolving list markers.
    /// </summary>
    public static class DocumentTraversal {
        /// <summary>
        /// Describes list information for a paragraph.
        /// </summary>
        public readonly struct ListInfo {
            public ListInfo(int level, bool ordered, int start, NumberFormatValues? format, string? text) {
                Level = level;
                Ordered = ordered;
                Start = start;
                NumberFormat = format;
                LevelText = text;
            }

            public int Level { get; }
            public bool Ordered { get; }
            public int Start { get; }
            public NumberFormatValues? NumberFormat { get; }
            public string? LevelText { get; }
        }

        /// <summary>
        /// Enumerates all sections within the document.
        /// </summary>
        public static IEnumerable<WordSection> EnumerateSections(WordDocument document) {
            return document?.Sections ?? Enumerable.Empty<WordSection>();
        }

        /// <summary>
        /// Resolves list information for the given paragraph.
        /// </summary>
        /// <param name="paragraph">Paragraph to inspect.</param>
        /// <returns>List info for the paragraph or null when paragraph isn't a list item.</returns>
        public static ListInfo? GetListInfo(WordParagraph paragraph) {
            if (paragraph == null || !paragraph.IsListItem) {
                return null;
            }

            int level = paragraph.ListItemLevel ?? 0;
            int start = 1;
            NumberFormatValues? numberFormat = null;
            string? levelText = null;

            int? numberId = paragraph._listNumberId;
            var list = numberId.HasValue ? paragraph._document?.Lists.FirstOrDefault(l => l._numberId == numberId) : null;
            var wordLevel = list?.Numbering.Levels.FirstOrDefault(l => l._level.LevelIndex == level);
            if (wordLevel != null) {
                start = wordLevel.StartNumberingValue;
                numberFormat = wordLevel._level.NumberingFormat?.Val;
                levelText = wordLevel.LevelText;
            }

            bool ordered = paragraph.ListStyle switch {
                WordListStyle.Bulleted => false,
                WordListStyle.BulletedChars => false,
                _ => true,
            };
            return new ListInfo(level, ordered, start, numberFormat, levelText);
        }

        /// <summary>
        /// Builds a lookup of list markers for all paragraphs in the document.
        /// </summary>
        public static Dictionary<WordParagraph, (int Level, string Marker)> BuildListMarkers(WordDocument document) {
            Dictionary<WordParagraph, (int, string)> result = new(ParagraphReferenceComparer.Instance);

            foreach (WordList list in document.Lists) {
                Dictionary<int, int> indices = new();
                Dictionary<int, NumberFormatValues?> formats = new();
                foreach (WordParagraph item in list.ListItems) {
                    ListInfo? info = GetListInfo(item);
                    if (info == null) {
                        continue;
                    }

                    int level = info.Value.Level;
                    if (!indices.ContainsKey(level)) {
                        indices[level] = info.Value.Start;
                        formats[level] = info.Value.NumberFormat;
                    }

                    int currentIndex = indices[level];
                    indices[level] = currentIndex + 1;

                    string marker = info.Value.Ordered
                        ? BuildMarker(level, currentIndex, indices, formats, info.Value.LevelText)
                        : (info.Value.LevelText ?? "•");

                    result[item] = (level, marker);
                }
            }

            return result;
        }

        private sealed class ParagraphReferenceComparer : IEqualityComparer<WordParagraph> {
            public static readonly ParagraphReferenceComparer Instance = new();
            public bool Equals(WordParagraph? x, WordParagraph? y) => ReferenceEquals(x, y);
            public int GetHashCode(WordParagraph obj) => RuntimeHelpers.GetHashCode(obj);
        }

        private static string BuildMarker(int level, int index, Dictionary<int, int> indices, Dictionary<int, NumberFormatValues?> formats, string? pattern) {
            if (string.IsNullOrEmpty(pattern)) {
                string formatted = FormatNumber(index, formats[level]);
                return formatted + ".";
            }

            string marker = pattern;
            marker = marker.Replace("%CurrentLevel", FormatNumber(index, formats[level]));
            marker = Regex.Replace(marker, "%([0-9]+)", m => {
                int lvl = int.Parse(m.Groups[1].Value) - 1;
                int value = lvl == level ? index : indices.TryGetValue(lvl, out int val) ? val - 1 : 0;
                formats.TryGetValue(lvl, out NumberFormatValues? fmt);
                return FormatNumber(value, fmt);
            });
            return marker;
        }

        private static string FormatNumber(int number, NumberFormatValues? format) {
            if (format == NumberFormatValues.LowerRoman) {
                return ToRoman(number).ToLowerInvariant();
            }
            if (format == NumberFormatValues.UpperRoman) {
                return ToRoman(number);
            }
            if (format == NumberFormatValues.LowerLetter) {
                return ((char)("a"[0] + number - 1)).ToString();
            }
            if (format == NumberFormatValues.UpperLetter) {
                return ((char)("A"[0] + number - 1)).ToString();
            }
            return number.ToString();
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

            StringBuilder sb = new();
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

