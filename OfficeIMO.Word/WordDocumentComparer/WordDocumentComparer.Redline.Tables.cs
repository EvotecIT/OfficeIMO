using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void ApplyTableFindings(WordprocessingDocument sourceDocument, WordprocessingDocument targetDocument, WordComparisonResult result, WordComparisonRedlineOptions options) {
            List<RedlineTableEntry> sourceTables = GetRedlineTableEntries(sourceDocument);
            List<RedlineTableEntry> targetTables = GetRedlineTableEntries(targetDocument);
            var rewrittenSourceTables = new HashSet<string>(StringComparer.Ordinal);
            var rewrittenTargetTables = new HashSet<string>(StringComparer.Ordinal);
            var insertedBeforeTableAnchors = new Dictionary<Table, OpenXmlElement>();
            var insertedAfterTableAnchors = new Dictionary<Table, OpenXmlElement>();
            var appendedDeletedTables = new Dictionary<string, OpenXmlElement>(StringComparer.Ordinal);

            foreach (WordComparisonFinding finding in result.Findings) {
                if (!ShouldTrackFinding(finding, options) ||
                    finding.Scope != WordComparisonScope.Table ||
                    !TryParseTableLocation(finding.Location, out int tableIndex) ||
                    tableIndex < 0 ||
                    !HasTrackedText(finding)) {
                    continue;
                }

                switch (finding.ChangeKind) {
                    case WordComparisonChangeKind.Inserted:
                        if (tableIndex < targetTables.Count) {
                            string targetKey = CreateRedlineTableRewriteKey(targetTables[tableIndex]);
                            if (rewrittenTargetTables.Contains(targetKey)) {
                                break;
                            }

                            RewriteTableWithTrackedText(targetTables[tableIndex].Table, trackInserted: true, options);
                            RemoveEmptyWordColorAttributes(targetTables[tableIndex].Table);
                            rewrittenTargetTables.Add(targetKey);
                        }

                        break;
                    case WordComparisonChangeKind.Deleted:
                        if (tableIndex < sourceTables.Count) {
                            string sourceKey = CreateRedlineTableRewriteKey(sourceTables[tableIndex]);
                            if (rewrittenSourceTables.Contains(sourceKey)) {
                                break;
                            }

                            InsertDeletedTable(targetDocument, sourceTables, targetTables, sourceTables[tableIndex], options, insertedBeforeTableAnchors, insertedAfterTableAnchors, appendedDeletedTables);
                            rewrittenSourceTables.Add(sourceKey);
                        }

                        break;
                }
            }
        }

        private static string CreateRedlineTableRewriteKey(RedlineTableEntry entry) {
            return entry.PartKey + "/" + entry.LocalIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static void RewriteTableWithTrackedText(Table table, bool trackInserted, WordComparisonRedlineOptions options) {
            foreach (TableRow row in table.Elements<TableRow>()) {
                RewriteRowWithTrackedText(row, trackInserted, options);
            }
        }

        private static void InsertDeletedTable(
            WordprocessingDocument targetDocument,
            IReadOnlyList<RedlineTableEntry> sourceTables,
            IReadOnlyList<RedlineTableEntry> targetTables,
            RedlineTableEntry sourceEntry,
            WordComparisonRedlineOptions options,
            Dictionary<Table, OpenXmlElement> insertedBeforeTableAnchors,
            Dictionary<Table, OpenXmlElement> insertedAfterTableAnchors,
            Dictionary<string, OpenXmlElement> appendedDeletedTables) {
            var deletedTable = (Table)sourceEntry.Table.CloneNode(true);
            RewriteTableWithTrackedText(deletedTable, trackInserted: false, options);
            RemoveEmptyWordColorAttributes(deletedTable);

            if (TryInsertDeletedNestedTable(sourceTables, targetTables, sourceEntry, deletedTable)) {
                return;
            }

            RedlineTableEntry? nextTable = FindNextSurvivingTargetTable(sourceTables, targetTables, sourceEntry) ??
                                           targetTables
                                               .Where(entry => string.Equals(entry.PartKey, sourceEntry.PartKey, StringComparison.Ordinal) && entry.LocalIndex >= sourceEntry.LocalIndex)
                                               .OrderBy(entry => entry.LocalIndex)
                                               .FirstOrDefault();
            if (nextTable != null) {
                InsertBeforeTableAnchor(nextTable.Table, deletedTable, insertedBeforeTableAnchors);
                return;
            }

            RedlineTableEntry? previousTable = targetTables
                .Where(entry => string.Equals(entry.PartKey, sourceEntry.PartKey, StringComparison.Ordinal) && entry.LocalIndex < sourceEntry.LocalIndex)
                .OrderByDescending(entry => entry.LocalIndex)
                .FirstOrDefault();
            if (previousTable != null) {
                InsertAfterTableAnchor(previousTable.Table, deletedTable, insertedAfterTableAnchors);
                return;
            }

            OpenXmlCompositeElement? targetContainer = GetRedlineContainerByPartKey(targetDocument, sourceEntry.PartKey);
            if (targetContainer == null) {
                return;
            }

            if (appendedDeletedTables.TryGetValue(sourceEntry.PartKey, out OpenXmlElement? previousAppended)) {
                previousAppended.InsertAfterSelf(deletedTable);
            } else {
                AppendRedlineTable(targetContainer, deletedTable);
            }

            appendedDeletedTables[sourceEntry.PartKey] = deletedTable;
        }

        private static RedlineTableEntry? FindNextSurvivingTargetTable(
            IReadOnlyList<RedlineTableEntry> sourceTables,
            IReadOnlyList<RedlineTableEntry> targetTables,
            RedlineTableEntry sourceEntry) {
            int sourceIndex = FindRedlineTableEntryIndex(sourceTables, sourceEntry.Table);
            if (sourceIndex < 0) {
                return null;
            }

            for (int followingIndex = sourceIndex + 1; followingIndex < sourceTables.Count; followingIndex++) {
                RedlineTableEntry followingSource = sourceTables[followingIndex];
                if (!string.Equals(followingSource.PartKey, sourceEntry.PartKey, StringComparison.Ordinal)) {
                    continue;
                }

                string followingIdentity = GetRedlineTableIdentity(followingSource.Table);
                RedlineTableEntry? targetMatch = targetTables
                    .Where(entry => string.Equals(entry.PartKey, sourceEntry.PartKey, StringComparison.Ordinal) &&
                                    string.Equals(GetRedlineTableIdentity(entry.Table), followingIdentity, StringComparison.Ordinal))
                    .OrderBy(entry => entry.LocalIndex)
                    .FirstOrDefault();
                if (targetMatch != null) {
                    return targetMatch;
                }
            }

            return null;
        }

        private static int FindRedlineTableEntryIndex(IReadOnlyList<RedlineTableEntry> entries, Table table) {
            for (int index = 0; index < entries.Count; index++) {
                if (ReferenceEquals(entries[index].Table, table)) {
                    return index;
                }
            }

            return -1;
        }

        private static string GetRedlineTableIdentity(Table table) {
            return string.Join("\u001e", table.Elements<TableRow>().Select(row =>
                string.Join("\u001f", row.Elements<TableCell>().Select(cell => cell.InnerText).ToArray())).ToArray());
        }

        private static string GetRedlineTableSurfaceIdentity(Table table) {
            return string.Join("\u001e", table.Elements<TableRow>().Select(row =>
                string.Join("\u001f", row.Elements<TableCell>().Select(cell => GetCellSurfaceText(table, cell)).ToArray())).ToArray());
        }

        private static string GetCellSurfaceText(Table ownerTable, TableCell cell) {
            return string.Concat(cell.Descendants<Text>()
                .Where(text => ReferenceEquals(text.Ancestors<Table>().FirstOrDefault(), ownerTable))
                .Select(text => text.Text));
        }

        private static bool IsNestedTable(Table table) {
            return table.Ancestors<TableCell>().Any();
        }

        private static void InsertBeforeTableAnchor(Table anchorTable, OpenXmlElement deletedTable, Dictionary<Table, OpenXmlElement> insertedBeforeTableAnchors) {
            if (insertedBeforeTableAnchors.TryGetValue(anchorTable, out OpenXmlElement? previousInserted)) {
                previousInserted.InsertAfterSelf(deletedTable);
            } else {
                anchorTable.InsertBeforeSelf(deletedTable);
            }

            insertedBeforeTableAnchors[anchorTable] = deletedTable;
        }

        private static void InsertAfterTableAnchor(Table anchorTable, OpenXmlElement deletedTable, Dictionary<Table, OpenXmlElement> insertedAfterTableAnchors) {
            if (insertedAfterTableAnchors.TryGetValue(anchorTable, out OpenXmlElement? previousInserted)) {
                previousInserted.InsertAfterSelf(deletedTable);
            } else {
                anchorTable.InsertAfterSelf(deletedTable);
            }

            insertedAfterTableAnchors[anchorTable] = deletedTable;
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
                string key = FootnotePartKeyPrefix + GetNotePartKeyId(footnotes[footnoteIndex], footnoteIndex);
                if (string.Equals(partKey, key, StringComparison.Ordinal)) {
                    return footnotes[footnoteIndex];
                }
            }

            List<Endnote> endnotes = GetReferencedEndnotes(mainPart);
            for (int endnoteIndex = 0; endnoteIndex < endnotes.Count; endnoteIndex++) {
                string key = EndnotePartKeyPrefix + GetNotePartKeyId(endnotes[endnoteIndex], endnoteIndex);
                if (string.Equals(partKey, key, StringComparison.Ordinal)) {
                    return endnotes[endnoteIndex];
                }
            }

            return null;
        }

        private static bool TryInsertDeletedNestedTable(IReadOnlyList<RedlineTableEntry> sourceTables, IReadOnlyList<RedlineTableEntry> targetTables, RedlineTableEntry sourceEntry, Table deletedTable) {
            if (!TryGetNestedTablePlacement(sourceTables, sourceEntry.Table, out NestedTablePlacement placement)) {
                return false;
            }

            RedlineTableEntry? targetParentEntry = FindTargetParentTableForNestedDeletion(sourceTables, targetTables, sourceEntry, placement);
            if (targetParentEntry == null) {
                return false;
            }

            Table targetParentTable = targetParentEntry.Table;
            List<TableRow> targetRows = targetParentTable.Elements<TableRow>().ToList();
            RedlineTableEntry sourceParentEntry = sourceTables[placement.ParentTableIndex];
            List<TableRow> sourceRows = sourceParentEntry.Table.Elements<TableRow>().ToList();
            int targetRowIndex = FindBestTargetBySimilarity(sourceRows, targetRows, placement.RowIndex, GetOpenXmlRowText, GetOpenXmlRowText);
            if (targetRowIndex < 0 || targetRowIndex >= targetRows.Count || placement.RowIndex < 0 || placement.RowIndex >= sourceRows.Count) {
                return false;
            }

            List<TableCell> sourceCells = sourceRows[placement.RowIndex].Elements<TableCell>().ToList();
            List<TableCell> targetCells = targetRows[targetRowIndex].Elements<TableCell>().ToList();
            int targetCellIndex = FindBestTargetBySimilarity(sourceCells, targetCells, placement.CellIndex, GetOpenXmlCellText, GetOpenXmlCellText);
            if (targetCellIndex < 0 || targetCellIndex >= targetCells.Count) {
                return false;
            }

            InsertNestedTableIntoCell(targetCells[targetCellIndex], deletedTable);
            return true;
        }

        private static int FindBestTargetBySimilarity<TSource, TTarget>(
            IReadOnlyList<TSource> sourceItems,
            IReadOnlyList<TTarget> targetItems,
            int sourceIndex,
            Func<TSource, string> sourceIdentity,
            Func<TTarget, string> targetIdentity) {
            if (sourceIndex < 0 || sourceIndex >= sourceItems.Count || targetItems.Count == 0) {
                return -1;
            }

            string sourceText = sourceIdentity(sourceItems[sourceIndex]);
            int bestIndex = -1;
            double bestSimilarity = double.MinValue;
            for (int index = 0; index < targetItems.Count; index++) {
                double similarity = GetTextSimilarity(sourceText, targetIdentity(targetItems[index]));
                if (similarity > bestSimilarity) {
                    bestSimilarity = similarity;
                    bestIndex = index;
                }
            }

            return bestIndex;
        }

        private static RedlineTableEntry? FindTargetParentTableForNestedDeletion(
            IReadOnlyList<RedlineTableEntry> sourceTables,
            IReadOnlyList<RedlineTableEntry> targetTables,
            RedlineTableEntry sourceEntry,
            NestedTablePlacement placement) {
            if (placement.ParentTableIndex < 0 || placement.ParentTableIndex >= sourceTables.Count) {
                return null;
            }

            RedlineTableEntry sourceParentEntry = sourceTables[placement.ParentTableIndex];
            string parentSurfaceIdentity = GetRedlineTableSurfaceIdentity(sourceParentEntry.Table);
            RedlineTableEntry? surfaceMatch = targetTables
                .Where(entry => string.Equals(entry.PartKey, sourceEntry.PartKey, StringComparison.Ordinal) &&
                                !IsNestedTable(entry.Table) &&
                                string.Equals(GetRedlineTableSurfaceIdentity(entry.Table), parentSurfaceIdentity, StringComparison.Ordinal))
                .OrderBy(entry => Math.Abs(entry.LocalIndex - sourceParentEntry.LocalIndex))
                .FirstOrDefault();
            if (surfaceMatch != null) {
                return surfaceMatch;
            }

            return targetTables
                .Where(entry => string.Equals(entry.PartKey, sourceEntry.PartKey, StringComparison.Ordinal) &&
                                !IsNestedTable(entry.Table) &&
                                entry.LocalIndex <= sourceParentEntry.LocalIndex)
                .OrderByDescending(entry => entry.LocalIndex)
                .FirstOrDefault();
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
