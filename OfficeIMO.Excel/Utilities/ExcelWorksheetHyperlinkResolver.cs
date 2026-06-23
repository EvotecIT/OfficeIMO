using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Utilities {
    internal static class ExcelWorksheetHyperlinkResolver {
        internal static Dictionary<string, ExcelHyperlinkSnapshot> BuildMap(WorksheetPart worksheetPart) {
            if (worksheetPart == null) {
                throw new ArgumentNullException(nameof(worksheetPart));
            }

            var worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
            var hyperlinks = worksheet.Elements<Hyperlinks>().FirstOrDefault();
            if (hyperlinks == null) {
                return new Dictionary<string, ExcelHyperlinkSnapshot>(StringComparer.OrdinalIgnoreCase);
            }

            var externalRelationships = worksheetPart.HyperlinkRelationships
                .Where(relationship => !string.IsNullOrWhiteSpace(relationship.Id))
                .ToDictionary(relationship => relationship.Id!, relationship => relationship, StringComparer.OrdinalIgnoreCase);
            var map = new Dictionary<string, ExcelHyperlinkSnapshot>(StringComparer.OrdinalIgnoreCase);

            foreach (Hyperlink hyperlink in hyperlinks.Elements<Hyperlink>()) {
                string? reference = hyperlink.Reference?.Value;
                if (string.IsNullOrWhiteSpace(reference)) {
                    continue;
                }

                ExcelHyperlinkSnapshot? snapshot = CreateSnapshot(hyperlink, externalRelationships);
                if (snapshot == null) {
                    continue;
                }

                AddReference(map, reference!, snapshot);
            }

            return map;
        }

        private static ExcelHyperlinkSnapshot? CreateSnapshot(Hyperlink hyperlink, IReadOnlyDictionary<string, HyperlinkRelationship> externalRelationships) {
            string? target = null;
            bool isExternal = false;

            string? relationshipId = hyperlink.Id?.Value;
            if (!string.IsNullOrWhiteSpace(relationshipId) && externalRelationships.TryGetValue(relationshipId!, out HyperlinkRelationship? relationship)) {
                target = relationship.Uri?.OriginalString;
                isExternal = true;
            } else if (!string.IsNullOrWhiteSpace(hyperlink.Location?.Value)) {
                target = hyperlink.Location!.Value!;
            }

            if (string.IsNullOrWhiteSpace(target)) {
                return null;
            }

            return new ExcelHyperlinkSnapshot {
                IsExternal = isExternal,
                Target = target!,
            };
        }

        private static void AddReference(Dictionary<string, ExcelHyperlinkSnapshot> map, string reference, ExcelHyperlinkSnapshot hyperlink) {
            string normalized = reference.Trim().Replace("$", string.Empty);
            if (A1.TryParseRange(normalized, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                for (int row = firstRow; row <= lastRow; row++) {
                    for (int column = firstColumn; column <= lastColumn; column++) {
                        map[A1.CellReference(row, column)] = hyperlink;
                    }
                }

                return;
            }

            (int singleRow, int singleColumn) = A1.ParseCellRef(normalized);
            if (singleRow > 0 && singleColumn > 0) {
                map[A1.CellReference(singleRow, singleColumn)] = hyperlink;
            }
        }
    }
}
