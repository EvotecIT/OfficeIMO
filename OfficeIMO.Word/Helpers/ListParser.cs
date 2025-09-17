using System.Collections.Generic;
using System.Linq;
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
            Dictionary<int, bool> listTypes = new();
            NumberingDefinitionsPart? numberingPart = mainPart.NumberingDefinitionsPart;
            if (numberingPart?.Numbering != null) {
                foreach (NumberingInstance instance in numberingPart.Numbering.Elements<NumberingInstance>()) {
                    var numberIdValue = instance.NumberID?.Value;
                    var abstractIdValue = instance.AbstractNumId?.Val?.Value;
                    if (numberIdValue == null || abstractIdValue == null) {
                        continue;
                    }

                    int id = numberIdValue.Value;
                    int absId = abstractIdValue.Value;
                    AbstractNum? abs = numberingPart.Numbering.Elements<AbstractNum>()
                        .FirstOrDefault(a => a.AbstractNumberId?.Value == absId);
                    bool ordered = true;
                    Level? lvl = abs?.Elements<Level>().FirstOrDefault(l => l.LevelIndex == 0);
                    NumberFormatValues? format = lvl?.NumberingFormat?.Val;
                    if (format == NumberFormatValues.Bullet) {
                        ordered = false;
                    }
                    listTypes[id] = ordered;
                }
            }
            return listTypes;
        }

        /// <summary>
        /// Determines whether a paragraph belongs to a bulleted list.
        /// </summary>
        /// <param name="paragraph">Paragraph to evaluate.</param>
        /// <param name="mainPart">Main document part.</param>
        public static bool IsBullet(Paragraph paragraph, MainDocumentPart mainPart) {
            var numProps = paragraph.ParagraphProperties?.NumberingProperties;
            if (numProps?.NumberingId?.Val?.Value == null) {
                return false;
            }
            int numId = numProps.NumberingId.Val.Value;
            var listTypes = GetListTypes(mainPart);
            return listTypes.TryGetValue(numId, out bool ordered) && !ordered;
        }

        /// <summary>
        /// Determines whether a paragraph belongs to an ordered (numbered) list.
        /// </summary>
        /// <param name="paragraph">Paragraph to evaluate.</param>
        /// <param name="mainPart">Main document part.</param>
        public static bool IsOrdered(Paragraph paragraph, MainDocumentPart mainPart) {
            var numProps = paragraph.ParagraphProperties?.NumberingProperties;
            if (numProps?.NumberingId?.Val?.Value == null) {
                return false;
            }
            int numId = numProps.NumberingId.Val.Value;
            var listTypes = GetListTypes(mainPart);
            return listTypes.TryGetValue(numId, out bool ordered) && ordered;
        }
    }
}
