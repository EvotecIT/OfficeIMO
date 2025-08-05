using OfficeIMO.Word;
using QuestPDF.Helpers;
using System;
using System.Collections.Generic;

namespace OfficeIMO.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static PageSize MapToPageSize(WordPageSize pageSize) {
            return pageSize switch {
                WordPageSize.Letter => PageSizes.Letter,
                WordPageSize.Legal => PageSizes.Legal,
                WordPageSize.Executive => PageSizes.Executive,
                WordPageSize.A3 => PageSizes.A3,
                WordPageSize.A4 => PageSizes.A4,
                WordPageSize.A5 => PageSizes.A5,
                WordPageSize.A6 => PageSizes.A6,
                WordPageSize.B5 => PageSizes.B5,
                _ => PageSizes.A4
            };
        }

        private static Dictionary<WordParagraph, (int Level, string Marker)> BuildListMarkers(WordDocument document) {
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
