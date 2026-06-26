using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal static class WordTableCellPropertiesNormalizer {
        internal static void Normalize(TableCellProperties tableCellProperties) {
            if (tableCellProperties.ChildElements.Count < 2) {
                return;
            }

            List<OpenXmlElement> ordered = tableCellProperties.ChildElements
                .Select((element, index) => new { Element = element, Index = index })
                .OrderBy(item => GetOrder(item.Element))
                .ThenBy(item => item.Index)
                .Select(item => item.Element)
                .ToList();

            for (int index = 0; index < ordered.Count; index++) {
                if (ReferenceEquals(tableCellProperties.ChildElements[index], ordered[index])) {
                    continue;
                }

                foreach (OpenXmlElement element in ordered) {
                    element.Remove();
                }

                foreach (OpenXmlElement element in ordered) {
                    tableCellProperties.AppendChild(element);
                }

                break;
            }
        }

        private static int GetOrder(OpenXmlElement element) {
            return element switch {
                ConditionalFormatStyle => 0,
                TableCellWidth => 1,
                GridSpan => 2,
                HorizontalMerge => 3,
                VerticalMerge => 4,
                TableCellBorders => 5,
                Shading => 6,
                NoWrap => 7,
                TableCellMargin => 8,
                TextDirection => 9,
                TableCellFitText => 10,
                TableCellVerticalAlignment => 11,
                HideMark => 12,
                CellInsertion => 13,
                CellDeletion => 14,
                CellMerge => 15,
                TableCellPropertiesChange => 16,
                _ => 100
            };
        }
    }
}
