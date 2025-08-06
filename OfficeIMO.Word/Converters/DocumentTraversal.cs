using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides helper methods for traversing documents and resolving list markers.
    /// </summary>
    public static class DocumentTraversal {
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
        /// <returns>Tuple containing list level and whether the list is ordered.</returns>
        public static (int Level, bool Ordered)? GetListInfo(WordParagraph paragraph) {
            if (paragraph == null || !paragraph.IsListItem) {
                return null;
            }

            int level = paragraph.ListItemLevel ?? 0;
            bool ordered = paragraph.ListStyle switch {
                WordListStyle.Bulleted => false,
                WordListStyle.BulletedChars => false,
                _ => true,
            };

            return (level, ordered);
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

