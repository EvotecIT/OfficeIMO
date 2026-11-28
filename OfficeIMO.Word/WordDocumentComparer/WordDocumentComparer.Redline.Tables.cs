using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void ApplyTableFindings(WordPackageView sourceDocument, WordPackageView targetDocument, WordComparisonResult result, WordComparisonRedlineOptions options) {
            List<RedlineTableEntry> sourceTables = GetRedlineTableEntries(sourceDocument);
            List<RedlineTableEntry> targetTables = GetRedlineTableEntries(targetDocument);
            var rewrittenSourceTables = new HashSet<string>(StringComparer.Ordinal);
            var rewrittenTargetTables = new HashSet<string>(StringComparer.Ordinal);
            var insertedBeforeTableAnchors = new Dictionary<OfficeXmlNode, OfficeXmlNode>();
            var insertedAfterTableAnchors = new Dictionary<OfficeXmlNode, OfficeXmlNode>();
            var appendedDeletedTables = new Dictionary<string, OfficeXmlNode>(StringComparer.Ordinal);

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

        private static void RewriteTableWithTrackedText(OfficeXmlNode table, bool trackInserted, WordComparisonRedlineOptions options) {
            foreach (OfficeXmlNode row in ChildrenW(table, "tr")) {
                RewriteRowWithTrackedText(row, trackInserted, options);
            }
        }

        private static void InsertDeletedTable(
            WordPackageView targetDocument,
            IReadOnlyList<RedlineTableEntry> sourceTables,
            IReadOnlyList<RedlineTableEntry> targetTables,
            RedlineTableEntry sourceEntry,
            WordComparisonRedlineOptions options,
            Dictionary<OfficeXmlNode, OfficeXmlNode> insertedBeforeTableAnchors,
            Dictionary<OfficeXmlNode, OfficeXmlNode> insertedAfterTableAnchors,
            Dictionary<string, OfficeXmlNode> appendedDeletedTables) {
            OfficeXmlNode deletedTable = sourceEntry.Table.Clone();
            WordPackageView.CopyGridNormalizationState(sourceEntry.Table, deletedTable);
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

            OfficeXmlNode targetContainer = GetRedlineContainerByPartKey(targetDocument, sourceEntry.PartKey);
            if (targetContainer.IsNone) {
                return;
            }

            if (appendedDeletedTables.TryGetValue(sourceEntry.PartKey, out OfficeXmlNode previousAppended)) {
                deletedTable = previousAppended.InsertAfterSelf(deletedTable);
            } else {
                deletedTable = AppendRedlineTable(targetContainer, deletedTable);
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

        private static int FindRedlineTableEntryIndex(IReadOnlyList<RedlineTableEntry> entries, OfficeXmlNode table) {
            for (int index = 0; index < entries.Count; index++) {
                if (entries[index].Table == table) {
                    return index;
                }
            }

            return -1;
        }

        private static string GetRedlineTableIdentity(OfficeXmlNode table) {
            return string.Join("\u001e", ChildrenW(table, "tr").Select(row =>
                string.Join("\u001f", ChildrenW(row, "tc").Select(InnerText).ToArray())).ToArray());
        }

        private static string GetRedlineTableSurfaceIdentity(OfficeXmlNode table) {
            return string.Join("\u001e", ChildrenW(table, "tr").Select(row =>
                string.Join("\u001f", ChildrenW(row, "tc").Select(cell => GetCellSurfaceText(table, cell)).ToArray())).ToArray());
        }

        private static string GetCellSurfaceText(OfficeXmlNode ownerTable, OfficeXmlNode cell) {
            return string.Concat(DescendantsW(cell, "t")
                .Where(text => AncestorW(text, "tbl") == ownerTable)
                .Select(text => text.InnerText));
        }

        private static bool IsNestedTable(OfficeXmlNode table) {
            return HasAncestorW(table, "tc");
        }

        private static void InsertBeforeTableAnchor(OfficeXmlNode anchorTable, OfficeXmlNode deletedTable, Dictionary<OfficeXmlNode, OfficeXmlNode> insertedBeforeTableAnchors) {
            if (insertedBeforeTableAnchors.TryGetValue(anchorTable, out OfficeXmlNode previousInserted)) {
                deletedTable = previousInserted.InsertAfterSelf(deletedTable);
            } else {
                deletedTable = anchorTable.Parent.InsertBefore(deletedTable, anchorTable);
            }

            insertedBeforeTableAnchors[anchorTable] = deletedTable;
        }

        private static void InsertAfterTableAnchor(OfficeXmlNode anchorTable, OfficeXmlNode deletedTable, Dictionary<OfficeXmlNode, OfficeXmlNode> insertedAfterTableAnchors) {
            if (insertedAfterTableAnchors.TryGetValue(anchorTable, out OfficeXmlNode previousInserted)) {
                deletedTable = previousInserted.InsertAfterSelf(deletedTable);
            } else {
                deletedTable = anchorTable.InsertAfterSelf(deletedTable);
            }

            insertedAfterTableAnchors[anchorTable] = deletedTable;
        }

        private static OfficeXmlNode AppendRedlineTable(OfficeXmlNode targetContainer, OfficeXmlNode deletedTable) {
            if (targetContainer.Is(W, "body")) {
                OfficeXmlNode sectionProperties = targetContainer.Children(W, "sectPr").Last();
                if (!sectionProperties.IsNone) {
                    return targetContainer.InsertBefore(deletedTable, sectionProperties);
                }
            }

            return targetContainer.Append(deletedTable);
        }

        private static OfficeXmlNode GetRedlineContainerByPartKey(WordPackageView document, string partKey) {
            WordPartView? mainPart = document.Main;
            if (string.Equals(partKey, BodyPartKey, StringComparison.Ordinal)) {
                return BodyOf(mainPart);
            }

            if (mainPart == null) {
                return default;
            }

            foreach (KeyValuePair<WordPartView, string> headerPartKey in CreateOrderedHeaderPartKeys(mainPart)) {
                if (string.Equals(partKey, headerPartKey.Value, StringComparison.Ordinal)) {
                    return HeaderRoot(headerPartKey.Key);
                }
            }

            foreach (KeyValuePair<WordPartView, string> footerPartKey in CreateOrderedFooterPartKeys(mainPart)) {
                if (string.Equals(partKey, footerPartKey.Value, StringComparison.Ordinal)) {
                    return FooterRoot(footerPartKey.Key);
                }
            }

            List<OfficeXmlNode> footnotes = GetReferencedFootnotes(mainPart);
            for (int footnoteIndex = 0; footnoteIndex < footnotes.Count; footnoteIndex++) {
                string key = FootnotePartKeyPrefix + GetNotePartKeyId(footnotes[footnoteIndex], footnoteIndex);
                if (string.Equals(partKey, key, StringComparison.Ordinal)) {
                    return footnotes[footnoteIndex];
                }
            }

            List<OfficeXmlNode> endnotes = GetReferencedEndnotes(mainPart);
            for (int endnoteIndex = 0; endnoteIndex < endnotes.Count; endnoteIndex++) {
                string key = EndnotePartKeyPrefix + GetNotePartKeyId(endnotes[endnoteIndex], endnoteIndex);
                if (string.Equals(partKey, key, StringComparison.Ordinal)) {
                    return endnotes[endnoteIndex];
                }
            }

            return default;
        }

        private static bool TryInsertDeletedNestedTable(IReadOnlyList<RedlineTableEntry> sourceTables, IReadOnlyList<RedlineTableEntry> targetTables, RedlineTableEntry sourceEntry, OfficeXmlNode deletedTable) {
            if (!TryGetNestedTablePlacement(sourceTables, sourceEntry.Table, out NestedTablePlacement placement)) {
                return false;
            }

            RedlineTableEntry? targetParentEntry = FindTargetParentTableForNestedDeletion(sourceTables, targetTables, sourceEntry, placement);
            if (targetParentEntry == null) {
                return false;
            }

            OfficeXmlNode targetParentTable = targetParentEntry.Table;
            List<OfficeXmlNode> targetRows = ChildrenW(targetParentTable, "tr");
            RedlineTableEntry sourceParentEntry = sourceTables[placement.ParentTableIndex];
            List<OfficeXmlNode> sourceRows = ChildrenW(sourceParentEntry.Table, "tr");
            int targetRowIndex = FindBestTargetBySimilarity(sourceRows, targetRows, placement.RowIndex, GetOpenXmlRowText, GetOpenXmlRowText);
            if (targetRowIndex < 0 || targetRowIndex >= targetRows.Count || placement.RowIndex < 0 || placement.RowIndex >= sourceRows.Count) {
                return false;
            }

            List<OfficeXmlNode> sourceCells = ChildrenW(sourceRows[placement.RowIndex], "tc");
            List<OfficeXmlNode> targetCells = ChildrenW(targetRows[targetRowIndex], "tc");
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

        private static bool TryGetNestedTablePlacement(IReadOnlyList<RedlineTableEntry> sourceTables, OfficeXmlNode sourceTable, out NestedTablePlacement placement) {
            placement = default;
            OfficeXmlNode parentCell = AncestorW(sourceTable, "tc");
            if (parentCell.IsNone) {
                return false;
            }

            OfficeXmlNode parentTable = AncestorW(parentCell, "tbl");
            OfficeXmlNode parentRow = AncestorW(parentCell, "tr");
            if (parentTable.IsNone || parentRow.IsNone) {
                return false;
            }

            int parentTableIndex = FindTableIndex(sourceTables, parentTable);
            if (parentTableIndex < 0) {
                return false;
            }

            List<OfficeXmlNode> rows = ChildrenW(parentTable, "tr");
            int rowIndex = rows.FindIndex(row => row == parentRow);
            if (rowIndex < 0) {
                return false;
            }

            List<OfficeXmlNode> cells = ChildrenW(parentRow, "tc");
            int cellIndex = cells.FindIndex(cell => cell == parentCell);
            if (cellIndex < 0) {
                return false;
            }

            placement = new NestedTablePlacement(parentTableIndex, rowIndex, cellIndex);
            return true;
        }

        private static int FindTableIndex(IReadOnlyList<RedlineTableEntry> tables, OfficeXmlNode table) {
            for (int index = 0; index < tables.Count; index++) {
                if (tables[index].Table == table) {
                    return index;
                }
            }

            return -1;
        }

        private static void InsertNestedTableIntoCell(OfficeXmlNode targetCell, OfficeXmlNode deletedTable) {
            OfficeXmlNode trailingParagraph = targetCell.Children(W, "p").Last();
            if (!trailingParagraph.IsNone && string.IsNullOrWhiteSpace(InnerText(trailingParagraph))) {
                targetCell.InsertBefore(deletedTable, trailingParagraph);
            } else {
                targetCell.Append(deletedTable);
            }

            if (!targetCell.LastChild.Is(W, "p")) {
                targetCell.Append(targetCell.CreateElement(W, "p"));
            }
        }

        private static void RemoveEmptyWordColorAttributes(OfficeXmlNode root) {
            RemoveEmptyWordColorAttribute(root);
            foreach (OfficeXmlNode element in Descendants(root)) {
                RemoveEmptyWordColorAttribute(element);
            }
        }

        private static void RemoveEmptyWordColorAttribute(OfficeXmlNode element) {
            string? color = element.Attribute(W, "color");
            if (color != null && string.IsNullOrWhiteSpace(color)) {
                element.RemoveAttribute(W, "color");
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
