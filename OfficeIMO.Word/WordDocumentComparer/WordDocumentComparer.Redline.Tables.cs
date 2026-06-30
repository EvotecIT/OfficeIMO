using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void ApplyTableFindings(WordprocessingDocument sourceDocument, WordprocessingDocument targetDocument, WordComparisonResult result, WordComparisonRedlineOptions options) {
            List<RedlineTableEntry> sourceTables = GetRedlineTableEntries(sourceDocument);
            List<RedlineTableEntry> targetTables = GetRedlineTableEntries(targetDocument);
            var rewrittenTables = new HashSet<int>();

            foreach (WordComparisonFinding finding in result.Findings) {
                if (!ShouldTrackFinding(finding, options) ||
                    finding.Scope != WordComparisonScope.Table ||
                    !TryParseTableLocation(finding.Location, out int tableIndex) ||
                    tableIndex < 0 ||
                    !HasTrackedText(finding)) {
                    continue;
                }

                if (rewrittenTables.Contains(tableIndex)) {
                    continue;
                }

                switch (finding.ChangeKind) {
                    case WordComparisonChangeKind.Inserted:
                        if (tableIndex < targetTables.Count) {
                            RewriteTableWithTrackedText(targetTables[tableIndex].Table, trackInserted: true, options);
                            RemoveEmptyWordColorAttributes(targetTables[tableIndex].Table);
                            rewrittenTables.Add(tableIndex);
                        }

                        break;
                    case WordComparisonChangeKind.Deleted:
                        if (tableIndex < sourceTables.Count) {
                            InsertDeletedTable(targetDocument, sourceTables, targetTables, sourceTables[tableIndex], tableIndex, options);
                            rewrittenTables.Add(tableIndex);
                        }

                        break;
                }
            }
        }

        private static void RewriteTableWithTrackedText(Table table, bool trackInserted, WordComparisonRedlineOptions options) {
            foreach (TableRow row in table.Elements<TableRow>()) {
                RewriteRowWithTrackedText(row, trackInserted, options);
            }
        }

        private static void InsertDeletedTable(WordprocessingDocument targetDocument, IReadOnlyList<RedlineTableEntry> sourceTables, IReadOnlyList<RedlineTableEntry> targetTables, RedlineTableEntry sourceEntry, int tableIndex, WordComparisonRedlineOptions options) {
            var deletedTable = (Table)sourceEntry.Table.CloneNode(true);
            RewriteTableWithTrackedText(deletedTable, trackInserted: false, options);
            RemoveEmptyWordColorAttributes(deletedTable);

            if (TryInsertDeletedNestedTable(sourceTables, targetTables, sourceEntry.Table, deletedTable)) {
                return;
            }

            if (tableIndex >= 0 && tableIndex < targetTables.Count) {
                targetTables[tableIndex].Table.InsertBeforeSelf(deletedTable);
                return;
            }

            OpenXmlCompositeElement? targetContainer = GetRedlineContainerByPartKey(targetDocument, sourceEntry.PartKey);
            if (targetContainer == null) {
                return;
            }

            AppendRedlineTable(targetContainer, deletedTable);
        }

        private static void AppendRedlineTable(OpenXmlCompositeElement targetContainer, Table deletedTable) {
            if (targetContainer is Body targetBody) {
                SectionProperties? sectionProperties = targetBody.Elements<SectionProperties>().LastOrDefault();
                if (sectionProperties != null) {
                    targetBody.InsertBefore(deletedTable, sectionProperties);
                    return;
                }
            }

            targetContainer.Append(deletedTable);
        }

        private static OpenXmlCompositeElement? GetRedlineContainerByPartKey(WordprocessingDocument document, string partKey) {
            MainDocumentPart? mainPart = document.MainDocumentPart;
            if (string.Equals(partKey, BodyPartKey, StringComparison.Ordinal)) {
                return mainPart?.Document?.Body;
            }

            if (mainPart == null) {
                return null;
            }

            foreach (KeyValuePair<HeaderPart, string> headerPartKey in CreateOrderedHeaderPartKeys(mainPart)) {
                if (string.Equals(partKey, headerPartKey.Value, StringComparison.Ordinal)) {
                    return headerPartKey.Key.Header;
                }
            }

            foreach (KeyValuePair<FooterPart, string> footerPartKey in CreateOrderedFooterPartKeys(mainPart)) {
                if (string.Equals(partKey, footerPartKey.Value, StringComparison.Ordinal)) {
                    return footerPartKey.Key.Footer;
                }
            }

            List<Footnote> footnotes = GetReferencedFootnotes(mainPart);
            for (int footnoteIndex = 0; footnoteIndex < footnotes.Count; footnoteIndex++) {
                string key = FootnotePartKeyPrefix + footnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (string.Equals(partKey, key, StringComparison.Ordinal)) {
                    return footnotes[footnoteIndex];
                }
            }

            List<Endnote> endnotes = GetReferencedEndnotes(mainPart);
            for (int endnoteIndex = 0; endnoteIndex < endnotes.Count; endnoteIndex++) {
                string key = EndnotePartKeyPrefix + endnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (string.Equals(partKey, key, StringComparison.Ordinal)) {
                    return endnotes[endnoteIndex];
                }
            }

            return null;
        }

        private static bool TryInsertDeletedNestedTable(IReadOnlyList<RedlineTableEntry> sourceTables, IReadOnlyList<RedlineTableEntry> targetTables, Table sourceTable, Table deletedTable) {
            if (!TryGetNestedTablePlacement(sourceTables, sourceTable, out NestedTablePlacement placement) ||
                placement.ParentTableIndex < 0 ||
                placement.ParentTableIndex >= targetTables.Count) {
                return false;
            }

            Table targetParentTable = targetTables[placement.ParentTableIndex].Table;
            List<TableRow> targetRows = targetParentTable.Elements<TableRow>().ToList();
            if (placement.RowIndex < 0 || placement.RowIndex >= targetRows.Count) {
                return false;
            }

            List<TableCell> targetCells = targetRows[placement.RowIndex].Elements<TableCell>().ToList();
            if (placement.CellIndex < 0 || placement.CellIndex >= targetCells.Count) {
                return false;
            }

            InsertNestedTableIntoCell(targetCells[placement.CellIndex], deletedTable);
            return true;
        }

        private static bool TryGetNestedTablePlacement(IReadOnlyList<RedlineTableEntry> sourceTables, Table sourceTable, out NestedTablePlacement placement) {
            placement = default;
            TableCell? parentCell = sourceTable.Ancestors<TableCell>().FirstOrDefault();
            if (parentCell == null) {
                return false;
            }

            Table? parentTable = parentCell.Ancestors<Table>().FirstOrDefault();
            TableRow? parentRow = parentCell.Ancestors<TableRow>().FirstOrDefault();
            if (parentTable == null || parentRow == null) {
                return false;
            }

            int parentTableIndex = FindTableIndex(sourceTables, parentTable);
            if (parentTableIndex < 0) {
                return false;
            }

            List<TableRow> rows = parentTable.Elements<TableRow>().ToList();
            int rowIndex = rows.FindIndex(row => ReferenceEquals(row, parentRow));
            if (rowIndex < 0) {
                return false;
            }

            List<TableCell> cells = parentRow.Elements<TableCell>().ToList();
            int cellIndex = cells.FindIndex(cell => ReferenceEquals(cell, parentCell));
            if (cellIndex < 0) {
                return false;
            }

            placement = new NestedTablePlacement(parentTableIndex, rowIndex, cellIndex);
            return true;
        }

        private static int FindTableIndex(IReadOnlyList<RedlineTableEntry> tables, Table table) {
            for (int index = 0; index < tables.Count; index++) {
                if (ReferenceEquals(tables[index].Table, table)) {
                    return index;
                }
            }

            return -1;
        }

        private static void InsertNestedTableIntoCell(TableCell targetCell, Table deletedTable) {
            Paragraph? trailingParagraph = targetCell.Elements<Paragraph>().LastOrDefault();
            if (trailingParagraph != null && string.IsNullOrWhiteSpace(trailingParagraph.InnerText)) {
                trailingParagraph.InsertBeforeSelf(deletedTable);
            } else {
                targetCell.Append(deletedTable);
            }

            if (targetCell.LastChild is not Paragraph) {
                targetCell.Append(new Paragraph());
            }
        }

        private static void RemoveEmptyWordColorAttributes(OpenXmlElement root) {
            foreach (OpenXmlElement element in new[] { root }.Concat(root.Descendants())) {
                foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                    if (attribute.LocalName == "color" &&
                        attribute.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main" &&
                        string.IsNullOrWhiteSpace(attribute.Value)) {
                        element.RemoveAttribute(attribute.LocalName, attribute.NamespaceUri);
                    }
                }
            }
        }

        private static bool TryParseTableLocation(string location, out int tableIndex) {
            tableIndex = -1;
            const string tablePrefix = "table[";
            if (!location.StartsWith(tablePrefix, StringComparison.Ordinal) || !location.EndsWith("]", StringComparison.Ordinal)) {
                return false;
            }

            if (location.IndexOf("]/", StringComparison.Ordinal) >= 0) {
                return false;
            }

            string tableText = location.Substring(tablePrefix.Length, location.Length - tablePrefix.Length - 1);
            return int.TryParse(tableText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out tableIndex);
        }

        private readonly struct NestedTablePlacement {
            internal NestedTablePlacement(int parentTableIndex, int rowIndex, int cellIndex) {
                ParentTableIndex = parentTableIndex;
                RowIndex = rowIndex;
                CellIndex = cellIndex;
            }

            internal int ParentTableIndex { get; }

            internal int RowIndex { get; }

            internal int CellIndex { get; }
        }
    }
}
