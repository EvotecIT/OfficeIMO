using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides helpers for detecting bullet and numbered lists within a document.
    /// </summary>
    public static class ListParser {
        /// <summary>
        /// Returns a dictionary describing list types keyed by numbering ID.
        /// </summary>
        /// <param name="mainPart">Main document part containing numbering definitions.</param>
        /// <returns>
        /// A dictionary where the key is a numbering ID and the value indicates whether the list is ordered (<c>true</c>)
        /// or bulleted (<c>false</c>).
        /// </returns>
        public static Dictionary<int, bool> GetListTypes(MainDocumentPart mainPart) {
            if (mainPart == null) throw new ArgumentNullException(nameof(mainPart));

            Dictionary<int, bool> listTypes = new();
            NumberingDefinitionsPart? numberingPart = mainPart.NumberingDefinitionsPart;
            if (numberingPart?.Numbering == null) {
                return listTypes;
            }

            var numbering = numberingPart.Numbering;
            foreach (NumberingInstance instance in numbering.Elements<NumberingInstance>()) {
                Int32Value? numberIdValue = instance.NumberID?.Value;
                int? numberId = numberIdValue?.Value;
                if (numberId is not int id) {
                    continue;
                }

                Int32Value? abstractIdValue = instance.AbstractNumId?.Val?.Value;
                int? abstractId = abstractIdValue?.Value;
                if (abstractId is not int absId) {
                    continue;
                }

                AbstractNum? abs = numbering.Elements<AbstractNum>()
                    .FirstOrDefault(a => {
                        Int32Value? abstractNumberIdValue = a.AbstractNumberId?.Value;
                        int? currentId = abstractNumberIdValue?.Value;
                        return currentId == absId;
                    });
                bool ordered = true;
                Level? lvl = abs?.Elements<Level>()
                    .FirstOrDefault(level => level.LevelIndex?.Value is int levelIndex && levelIndex == 0);
                EnumValue<NumberFormatValues>? format = lvl?.NumberingFormat?.Val;
                if (format?.Value == NumberFormatValues.Bullet) {
                    ordered = false;
                }
                listTypes[id] = ordered;
            }

            return listTypes;
        }

        /// <summary>
        /// Determines whether a paragraph belongs to a bulleted list.
        /// </summary>
        /// <param name="paragraph">Paragraph to evaluate.</param>
        /// <param name="mainPart">Main document part.</param>
        public static bool IsBullet(Paragraph paragraph, MainDocumentPart mainPart) {
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
            if (mainPart == null) throw new ArgumentNullException(nameof(mainPart));

            var numProps = paragraph.ParagraphProperties?.NumberingProperties;
            Int32Value? numberingIdValue = numProps?.NumberingId?.Val?.Value;
            int? numberId = numberingIdValue?.Value;
            if (numberId is not int numId) {
                return false;
            }

            var listTypes = GetListTypes(mainPart);
            return listTypes.TryGetValue(numId, out bool ordered) && !ordered;
        }

        /// <summary>
        /// Determines whether a paragraph belongs to an ordered (numbered) list.
        /// </summary>
        /// <param name="paragraph">Paragraph to evaluate.</param>
        /// <param name="mainPart">Main document part.</param>
        public static bool IsOrdered(Paragraph paragraph, MainDocumentPart mainPart) {
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
            if (mainPart == null) throw new ArgumentNullException(nameof(mainPart));
        
            var numProps = paragraph.ParagraphProperties?.NumberingProperties;
            Int32Value? numberingIdValue = numProps?.NumberingId?.Val?.Value;
            int? numberId = numberingIdValue?.Value;
            if (numberId is not int numId) {
                return false;
            }

            var listTypes = GetListTypes(mainPart);
            return listTypes.TryGetValue(numId, out bool ordered) && ordered;
        }
    }
}
