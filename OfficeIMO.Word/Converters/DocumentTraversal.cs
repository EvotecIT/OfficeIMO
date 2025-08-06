using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

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
            Dictionary<WordParagraph, (int, string)> result = new();

            foreach (WordList list in document.Lists) {
                Dictionary<int, int> indices = new();
                bool bullet = list.Style.ToString().IndexOf("Bullet", StringComparison.OrdinalIgnoreCase) >= 0;
                foreach (WordParagraph item in list.ListItems) {
                    int level = item.ListItemLevel ?? 0;
                    if (!indices.ContainsKey(level)) {
                        indices[level] = 1;
                    }

                    int index = indices[level];
                    indices[level] = index + 1;
                    string marker = bullet ? "•" : $"{index}.";
                    result[item] = (level, marker);
                }
            }

            return result;
        }
    }
}

