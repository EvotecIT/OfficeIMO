using OfficeIMO.Diff;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private const int LcsCellLimit = 1_000_000;
        private const int MaxComparisonAlignmentWindow = 256;
        private const long MaxComparisonTextWorkUnits = 64_000_000;
        private const long MaxComparisonDisclosureWorkUnits = 64_000_000;
        private const int MaxBoundedTextSimilaritySamples = 64;
        private const int MaxBoundedTextSimilarityAnchors = 8;
        private const int BoundedTextSimilarityAnchorLength = 8;
        private const int BodyPartOrderBase = 0;
        private const int HeaderPartOrderBase = 1_000_000;
        private const int FooterPartOrderBase = 2_000_000;
        private const int FootnotePartOrderBase = 3_000_000;
        private const int EndnotePartOrderBase = 4_000_000;
        private const int RelatedPartOrderStride = 100_000;
        private const string BodyPartKey = "body";
        private const string HeaderPartKeyPrefix = "header:";
        private const string FooterPartKeyPrefix = "footer:";
        private const string FootnotePartKeyPrefix = "footnote:";
        private const string EndnotePartKeyPrefix = "endnote:";
        private const string TableRowSeparator = "\n";
        private const string CellParagraphSeparator = "[ParagraphBreak]";

        /// <summary>
        /// Compares two documents and returns a machine-readable summary of structural differences.
        /// </summary>
        /// <param name="sourcePath">Path to the original document.</param>
        /// <param name="targetPath">Path to the modified document.</param>
        /// <returns>A deterministic comparison result that can be used for review reports and automation.</returns>
        public static WordComparisonResult CompareStructure(string sourcePath, string targetPath) {
            return CompareStructure(sourcePath, targetPath, null);
        }

        /// <summary>
        /// Compares two documents and returns a machine-readable summary of structural differences.
        /// </summary>
        /// <param name="sourcePath">Path to the original document.</param>
        /// <param name="targetPath">Path to the modified document.</param>
        /// <param name="options">Optional comparison switches for text normalization and feature families.</param>
        /// <returns>A deterministic comparison result that can be used for review reports and automation.</returns>
        public static WordComparisonResult CompareStructure(string sourcePath, string targetPath, WordComparisonOptions? options) {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (string.IsNullOrEmpty(targetPath)) throw new ArgumentNullException(nameof(targetPath));

            using WordDocument source = WordDocument.Load(sourcePath);
            using WordDocument target = WordDocument.Load(targetPath);
            return CompareStructureCore(source, target, sourcePath, targetPath, options);
        }

        /// <summary>
        /// Compares two already-loaded documents without writing them to disk and returns a machine-readable summary.
        /// </summary>
        /// <param name="source">Original document.</param>
        /// <param name="target">Modified document.</param>
        /// <param name="options">Optional comparison switches for text normalization and feature families.</param>
        /// <param name="sourceLabel">Logical source label included in reports.</param>
        /// <param name="targetLabel">Logical target label included in reports.</param>
        /// <returns>A deterministic comparison result that can be used for review reports and automation.</returns>
        public static WordComparisonResult CompareStructure(
            WordDocument source,
            WordDocument target,
            WordComparisonOptions? options = null,
            string sourceLabel = "source",
            string targetLabel = "target") {
            ArgumentNullException.ThrowIfNull(source);
            ArgumentNullException.ThrowIfNull(target);
            return CompareStructureCore(source, target, sourceLabel, targetLabel, options);
        }

        private static WordComparisonResult CompareStructureCore(
            WordDocument sourceDocument,
            WordDocument targetDocument,
            string sourceLabel,
            string targetLabel,
            WordComparisonOptions? options) {
            options ??= WordComparisonOptions.Default;

            ComparedDocument source = ComparedDocument.Open(sourceDocument, options);
            ComparedDocument target = ComparedDocument.Open(targetDocument, options);
            WordComparisonResult result = new(sourceLabel, targetLabel);
            var numberingStyleCatalogs = new ParagraphNumberingStyleCatalogCache();
            AnalyzeKnownComparisonLimitations(
                source.Package,
                target.Package,
                result,
                options,
                new ComparisonWorkBudget(MaxComparisonDisclosureWorkUnits),
                numberingStyleCatalogs);
            var comparisonWorkBudget = new ComparisonWorkBudget(MaxComparisonTextWorkUnits);
            AnalyzeParagraphs(source.Package, target.Package, result, options, comparisonWorkBudget);
            if (options.CompareFields) {
                AnalyzeFields(source, target, result, options);
            }

            if (options.CompareContentControls) {
                AnalyzeContentControls(source.Package, target.Package, result, options);
            }

            if (options.CompareBookmarks) {
                AnalyzeBookmarks(source.Package, target.Package, result, options);
            }

            if (options.CompareHyperlinks) {
                AnalyzeHyperlinks(source.Package, target.Package, result, options);
            }

            if (options.CompareLists) {
                AnalyzeLists(source.Package, target.Package, result, options, numberingStyleCatalogs);
            }

            if (options.CompareComments) {
                AnalyzeComments(source, target, result, options);
            }

            if (options.CompareRevisions) {
                AnalyzeRevisions(source, target, result, options);
            }

            AnalyzeTables(source.Package, target.Package, result, options, comparisonWorkBudget);
            if (options.CompareImages) {
                AnalyzeImages(source.Package, target.Package, result);
            }

            if (options.CompareShapes) {
                AnalyzeShapes(source.Package, target.Package, result, options);
            }

            if (options.CompareBlockOrder) {
                AnalyzeBlockOrder(source.Package, target.Package, result, options);
            }

            ApplyScopeFilters(result, options);
            result.SortFindingsByDocumentOrder();
            return result;
        }

        /// <summary>
        /// A document being compared and its package as the comparer reads it. A document that works without the SDK
        /// keeps doing so unless it holds fields or review markup, which the engine's field and review inventories read
        /// through the SDK; such a document opens it before any part is read, so every element read comes from one layer.
        /// </summary>
        private sealed class ComparedDocument {
            private ComparedDocument(WordDocument document, WordPackageView package, bool inspectsFields, bool inspectsReview) {
                Document = document;
                Package = package;
                InspectsFields = inspectsFields;
                InspectsReview = inspectsReview;
            }

            internal WordDocument Document { get; }

            internal WordPackageView Package { get; }

            /// <summary>Whether the document may hold fields; one that can't has none to inspect.</summary>
            internal bool InspectsFields { get; }

            /// <summary>Whether the document may hold comments or tracked revisions; one that can't has none to inspect.</summary>
            internal bool InspectsReview { get; }

            internal static ComparedDocument Open(WordDocument document, WordComparisonOptions options) {
                bool inspectsFields = true;
                bool inspectsReview = true;
                if (document.IsSdkDeferred) {
                    WordPackageView owned = WordPackageView.For(document);
                    inspectsFields = options.CompareFields && HasFieldMarkup(owned);
                    inspectsReview = (options.CompareComments || options.CompareRevisions) && HasReviewMarkup(owned);
                    if (inspectsFields || inspectsReview) {
                        document.OpenSdkPackage(inspectsFields ? "Document comparison: fields" : "Document comparison: review markup");
                    }
                }

                return new ComparedDocument(document, WordPackageView.For(document), inspectsFields, inspectsReview);
            }

            private static bool HasFieldMarkup(WordPackageView package) =>
                package.Main is { } mainPart && WordFieldInventory.EnumerateFieldRoots(mainPart).Any(root =>
                    Descendants(root.Root).Any(element => element.NamespaceId == W &&
                        element.LocalName is "fldSimple" or "fldChar" or "instrText" or "delInstrText"));

            private static bool HasReviewMarkup(WordPackageView package) {
                if (package.Main is not { } mainPart) return false;
                if (mainPart.FirstPart(WordPackageView.CommentsRelationshipType) != null) return true;
                return WordFieldInventory.EnumerateFieldRoots(mainPart).Any(root =>
                    // The review inventory reads revision markup by local name, in any namespace.
                    Descendants(root.Root).Any(element => IsReviewMarkupName(element.LocalName)));
            }

            private static bool IsReviewMarkupName(string localName) =>
                localName is "ins" or "del" or "moveFrom" or "moveTo" or "cellIns" or "cellDel" or "cellMerge"
                    or "moveFromRangeStart" or "moveToRangeStart" or "commentRangeStart" or "commentReference"
                || localName.EndsWith("Change", StringComparison.Ordinal);
        }

        private static void ApplyScopeFilters(WordComparisonResult result, WordComparisonOptions options) {
            bool hasIncludedScopes = options.IncludedScopes != null && options.IncludedScopes.Count > 0;
            bool hasExcludedScopes = options.ExcludedScopes != null && options.ExcludedScopes.Count > 0;
            if (!hasIncludedScopes && !hasExcludedScopes) {
                return;
            }

            result.RemoveWhere(finding =>
                (hasIncludedScopes && !options.IncludedScopes!.Contains(finding.Scope)) ||
                (hasExcludedScopes && options.ExcludedScopes!.Contains(finding.Scope)));
        }

        private static void AnalyzeTables(
            WordPackageView source,
            WordPackageView target,
            WordComparisonResult result,
            WordComparisonOptions options,
            ComparisonWorkBudget comparisonWorkBudget) {
            List<TableSnapshot> sourceTables = GetTableSnapshots(source, options);
            List<TableSnapshot> targetTables = GetTableSnapshots(target, options);
            IReadOnlyList<MatchedIndexPair> matchedTables = FindMatchingIndexes(
                sourceTables.Select(table => table.MatchKey).ToList(),
                targetTables.Select(table => table.MatchKey).ToList(),
                StringComparer.Ordinal);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedTables) {
                AddTableRangeFindings(sourceTables, targetTables, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result, options, comparisonWorkBudget);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddTableRangeFindings(sourceTables, targetTables, sourceStart, sourceTables.Count, targetStart, targetTables.Count, result, options, comparisonWorkBudget);
        }

        private static void AddTableRangeFindings(
            IReadOnlyList<TableSnapshot> sourceTables,
            IReadOnlyList<TableSnapshot> targetTables,
            int sourceStart,
            int sourceEnd,
            int targetStart,
            int targetEnd,
            WordComparisonResult result,
            WordComparisonOptions options,
            ComparisonWorkBudget comparisonWorkBudget) {
            int sourceIndex = sourceStart;
            int targetIndex = targetStart;

            while (sourceIndex < sourceEnd && targetIndex < targetEnd) {
                if (targetEnd - targetIndex > sourceEnd - sourceIndex &&
                    targetIndex + 1 < targetEnd &&
                    GetTableSimilarity(sourceTables[sourceIndex], targetTables[targetIndex + 1], comparisonWorkBudget) >
                    GetTableSimilarity(sourceTables[sourceIndex], targetTables[targetIndex], comparisonWorkBudget)) {
                    AddInsertedTableFinding(targetTables, targetIndex, result);
                    targetIndex++;
                    continue;
                }

                if (sourceEnd - sourceIndex > targetEnd - targetIndex &&
                    sourceIndex + 1 < sourceEnd &&
                    GetTableSimilarity(sourceTables[sourceIndex + 1], targetTables[targetIndex], comparisonWorkBudget) >
                    GetTableSimilarity(sourceTables[sourceIndex], targetTables[targetIndex], comparisonWorkBudget)) {
                    AddDeletedTableFinding(sourceTables, sourceIndex, result);
                    sourceIndex++;
                    continue;
                }

                if (!string.Equals(sourceTables[sourceIndex].PartKey, targetTables[targetIndex].PartKey, StringComparison.Ordinal)) {
                    AddDeletedTableFinding(sourceTables, sourceIndex, result);
                    AddInsertedTableFinding(targetTables, targetIndex, result);
                    sourceIndex++;
                    targetIndex++;
                    continue;
                }

                AnalyzeTable(sourceTables[sourceIndex], targetTables[targetIndex], targetIndex, targetTables[targetIndex].DocumentOrder, result, options, comparisonWorkBudget);
                sourceIndex++;
                targetIndex++;
            }

            while (targetIndex < targetEnd) {
                AddInsertedTableFinding(targetTables, targetIndex, result);
                targetIndex++;
            }

            while (sourceIndex < sourceEnd) {
                AddDeletedTableFinding(sourceTables, sourceIndex, result);
                sourceIndex++;
            }
        }

        private static void AddInsertedTableFinding(IReadOnlyList<TableSnapshot> targetTables, int tableIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.Table,
                WordComparisonChangeKind.Inserted,
                TableLocation(tableIndex),
                null,
                tableIndex,
                null,
                targetTables[tableIndex].Text,
                "Table inserted."),
                targetTables[tableIndex].DocumentOrder);
        }

        private static void AddDeletedTableFinding(IReadOnlyList<TableSnapshot> sourceTables, int tableIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.Table,
                WordComparisonChangeKind.Deleted,
                TableLocation(tableIndex),
                tableIndex,
                null,
                sourceTables[tableIndex].Text,
                null,
                "Table deleted."),
                sourceTables[tableIndex].DocumentOrder);
        }

        private static void AnalyzeTable(
            TableSnapshot source,
            TableSnapshot target,
            int tableIndex,
            int tableDocumentOrder,
            WordComparisonResult result,
            WordComparisonOptions options,
            ComparisonWorkBudget comparisonWorkBudget) {
            List<OfficeXmlNode> sourceRows = TableRows(source.Table);
            List<OfficeXmlNode> targetRows = TableRows(target.Table);
            List<string> sourceRowComparisonTexts = sourceRows
                .Select(row => NormalizeComparisonText(GetRowText(row), options))
                .ToList();
            List<string> targetRowComparisonTexts = targetRows
                .Select(row => NormalizeComparisonText(GetRowText(row), options))
                .ToList();
            IReadOnlyList<MatchedIndexPair> matchedRows = FindMatchingIndexes(
                sourceRows.Select(row => GetRowMatchKey(row, source.Part, options)).ToList(),
                targetRows.Select(row => GetRowMatchKey(row, target.Part, options)).ToList(),
                StringComparer.Ordinal);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedRows) {
                AddTableRowRangeFindings(sourceRows, targetRows, sourceRowComparisonTexts, targetRowComparisonTexts, source.Part, target.Part, tableIndex, tableDocumentOrder, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result, options, comparisonWorkBudget);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddTableRowRangeFindings(sourceRows, targetRows, sourceRowComparisonTexts, targetRowComparisonTexts, source.Part, target.Part, tableIndex, tableDocumentOrder, sourceStart, sourceRows.Count, targetStart, targetRows.Count, result, options, comparisonWorkBudget);
        }

        private static void AddTableRowRangeFindings(
            IReadOnlyList<OfficeXmlNode> sourceRows,
            IReadOnlyList<OfficeXmlNode> targetRows,
            IReadOnlyList<string> sourceRowComparisonTexts,
            IReadOnlyList<string> targetRowComparisonTexts,
            WordPartView? sourcePart,
            WordPartView? targetPart,
            int tableIndex,
            int tableDocumentOrder,
            int sourceStart,
            int sourceEnd,
            int targetStart,
            int targetEnd,
            WordComparisonResult result,
            WordComparisonOptions options,
            ComparisonWorkBudget comparisonWorkBudget) {
            int sourceIndex = sourceStart;
            int targetIndex = targetStart;

            while (sourceIndex < sourceEnd && targetIndex < targetEnd) {
                int betterTargetIndex = FindBetterTargetRowAlignmentIndex(
                    sourceRows[sourceIndex],
                    sourceRowComparisonTexts[sourceIndex],
                    targetRows,
                    targetRowComparisonTexts,
                    targetIndex,
                    targetEnd,
                    comparisonWorkBudget);
                if (betterTargetIndex > targetIndex) {
                    while (targetIndex < betterTargetIndex) {
                        AddInsertedTableRowFinding(targetRows, tableIndex, tableDocumentOrder, targetIndex, result);
                        targetIndex++;
                    }

                    continue;
                }

                int betterSourceIndex = FindBetterSourceRowAlignmentIndex(
                    sourceRows,
                    sourceRowComparisonTexts,
                    sourceIndex,
                    sourceEnd,
                    targetRows[targetIndex],
                    targetRowComparisonTexts[targetIndex],
                    comparisonWorkBudget);
                if (betterSourceIndex > sourceIndex) {
                    while (sourceIndex < betterSourceIndex) {
                        AddDeletedTableRowFinding(sourceRows, tableIndex, tableDocumentOrder, sourceIndex, result);
                        sourceIndex++;
                    }

                    continue;
                }

                if (RowCells(sourceRows[sourceIndex]).Count != RowCells(targetRows[targetIndex]).Count &&
                    string.Equals(sourceRowComparisonTexts[sourceIndex], targetRowComparisonTexts[targetIndex], StringComparison.Ordinal)) {
                    AddDeletedTableRowFinding(sourceRows, tableIndex, tableDocumentOrder, sourceIndex, result);
                    AddInsertedTableRowFinding(targetRows, tableIndex, tableDocumentOrder, targetIndex, result);
                    sourceIndex++;
                    targetIndex++;
                    continue;
                }

                AnalyzeTableRow(sourceRows[sourceIndex], targetRows[targetIndex], sourcePart, targetPart, tableIndex, tableDocumentOrder, sourceIndex, targetIndex, result, options, comparisonWorkBudget);
                sourceIndex++;
                targetIndex++;
            }

            while (targetIndex < targetEnd) {
                AddInsertedTableRowFinding(targetRows, tableIndex, tableDocumentOrder, targetIndex, result);
                targetIndex++;
            }

            while (sourceIndex < sourceEnd) {
                AddDeletedTableRowFinding(sourceRows, tableIndex, tableDocumentOrder, sourceIndex, result);
                sourceIndex++;
            }
        }

        private static int FindBetterTargetRowAlignmentIndex(
            OfficeXmlNode sourceRow,
            string sourceComparisonText,
            IReadOnlyList<OfficeXmlNode> targetRows,
            IReadOnlyList<string> targetComparisonTexts,
            int targetStart,
            int targetEnd,
            ComparisonWorkBudget comparisonWorkBudget) {
            double currentSimilarity = GetRowSimilarity(sourceRow, sourceComparisonText, targetRows[targetStart], targetComparisonTexts[targetStart], comparisonWorkBudget);
            int bestIndex = targetStart;
            double bestSimilarity = currentSimilarity;

            foreach (int index in SelectBoundedAlignmentCandidates(
                targetStart + 1,
                targetEnd,
                index => GetRowAlignmentPrefilterSimilarity(sourceRow, sourceComparisonText, targetRows[index], targetComparisonTexts[index], comparisonWorkBudget),
                comparisonWorkBudget)) {
                double similarity = GetRowSimilarity(sourceRow, sourceComparisonText, targetRows[index], targetComparisonTexts[index], comparisonWorkBudget);
                if (similarity > bestSimilarity) {
                    bestSimilarity = similarity;
                    bestIndex = index;
                }
            }

            return bestSimilarity > currentSimilarity ? bestIndex : targetStart;
        }

        private static int FindBetterSourceRowAlignmentIndex(
            IReadOnlyList<OfficeXmlNode> sourceRows,
            IReadOnlyList<string> sourceComparisonTexts,
            int sourceStart,
            int sourceEnd,
            OfficeXmlNode targetRow,
            string targetComparisonText,
            ComparisonWorkBudget comparisonWorkBudget) {
            double currentSimilarity = GetRowSimilarity(sourceRows[sourceStart], sourceComparisonTexts[sourceStart], targetRow, targetComparisonText, comparisonWorkBudget);
            int bestIndex = sourceStart;
            double bestSimilarity = currentSimilarity;

            foreach (int index in SelectBoundedAlignmentCandidates(
                sourceStart + 1,
                sourceEnd,
                index => GetRowAlignmentPrefilterSimilarity(sourceRows[index], sourceComparisonTexts[index], targetRow, targetComparisonText, comparisonWorkBudget),
                comparisonWorkBudget)) {
                double similarity = GetRowSimilarity(sourceRows[index], sourceComparisonTexts[index], targetRow, targetComparisonText, comparisonWorkBudget);
                if (similarity > bestSimilarity) {
                    bestSimilarity = similarity;
                    bestIndex = index;
                }
            }

            return bestSimilarity > currentSimilarity ? bestIndex : sourceStart;
        }

        private static void AddInsertedTableRowFinding(IReadOnlyList<OfficeXmlNode> targetRows, int tableIndex, int tableDocumentOrder, int rowIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.TableRow,
                WordComparisonChangeKind.Inserted,
                RowLocation(tableIndex, rowIndex),
                null,
                rowIndex,
                null,
                GetRowText(targetRows[rowIndex]),
                "Table row inserted."),
                GetTableChildDocumentOrder(tableDocumentOrder, targetRows[rowIndex]));
        }

        private static void AddDeletedTableRowFinding(IReadOnlyList<OfficeXmlNode> sourceRows, int tableIndex, int tableDocumentOrder, int rowIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.TableRow,
                WordComparisonChangeKind.Deleted,
                RowLocation(tableIndex, rowIndex),
                rowIndex,
                null,
                GetRowText(sourceRows[rowIndex]),
                null,
                "Table row deleted."),
                GetTableChildDocumentOrder(tableDocumentOrder, sourceRows[rowIndex]));
        }

        private static void AnalyzeTableRow(
            OfficeXmlNode source,
            OfficeXmlNode target,
            WordPartView? sourcePart,
            WordPartView? targetPart,
            int tableIndex,
            int tableDocumentOrder,
            int sourceRowIndex,
            int targetRowIndex,
            WordComparisonResult result,
            WordComparisonOptions options,
            ComparisonWorkBudget comparisonWorkBudget) {
            List<OfficeXmlNode> sourceCells = RowCells(source);
            List<OfficeXmlNode> targetCells = RowCells(target);
            List<string> sourceCellComparisonTexts = sourceCells
                .Select(cell => NormalizeComparisonText(GetCellText(cell), options))
                .ToList();
            List<string> targetCellComparisonTexts = targetCells
                .Select(cell => NormalizeComparisonText(GetCellText(cell), options))
                .ToList();
            IReadOnlyList<MatchedIndexPair> matchedCells = FindMatchingIndexes(
                sourceCells.Select(cell => GetCellMatchKey(cell, sourcePart, options)).ToList(),
                targetCells.Select(cell => GetCellMatchKey(cell, targetPart, options)).ToList(),
                StringComparer.Ordinal);

            int sourceStart = 0;
            int targetStart = 0;

            foreach (MatchedIndexPair match in matchedCells) {
                AddTableCellRangeFindings(sourceCells, targetCells, sourceCellComparisonTexts, targetCellComparisonTexts, sourcePart, targetPart, tableIndex, tableDocumentOrder, sourceRowIndex, targetRowIndex, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result, options, comparisonWorkBudget);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddTableCellRangeFindings(sourceCells, targetCells, sourceCellComparisonTexts, targetCellComparisonTexts, sourcePart, targetPart, tableIndex, tableDocumentOrder, sourceRowIndex, targetRowIndex, sourceStart, sourceCells.Count, targetStart, targetCells.Count, result, options, comparisonWorkBudget);
        }

        private static void AddTableCellRangeFindings(
            IReadOnlyList<OfficeXmlNode> sourceCells,
            IReadOnlyList<OfficeXmlNode> targetCells,
            IReadOnlyList<string> sourceCellComparisonTexts,
            IReadOnlyList<string> targetCellComparisonTexts,
            WordPartView? sourcePart,
            WordPartView? targetPart,
            int tableIndex,
            int tableDocumentOrder,
            int sourceRowIndex,
            int targetRowIndex,
            int sourceStart,
            int sourceEnd,
            int targetStart,
            int targetEnd,
            WordComparisonResult result,
            WordComparisonOptions options,
            ComparisonWorkBudget comparisonWorkBudget) {
            int sourceIndex = sourceStart;
            int targetIndex = targetStart;

            while (sourceIndex < sourceEnd && targetIndex < targetEnd) {
                if (targetEnd - targetIndex > sourceEnd - sourceIndex &&
                    targetIndex + 1 < targetEnd &&
                    GetCellSimilarity(sourceCellComparisonTexts[sourceIndex], targetCellComparisonTexts[targetIndex + 1], comparisonWorkBudget) >
                    GetCellSimilarity(sourceCellComparisonTexts[sourceIndex], targetCellComparisonTexts[targetIndex], comparisonWorkBudget)) {
                    AddInsertedTableCellFinding(targetCells, tableIndex, tableDocumentOrder, targetRowIndex, targetIndex, result);
                    targetIndex++;
                    continue;
                }

                if (sourceEnd - sourceIndex > targetEnd - targetIndex &&
                    sourceIndex + 1 < sourceEnd &&
                    GetCellSimilarity(sourceCellComparisonTexts[sourceIndex + 1], targetCellComparisonTexts[targetIndex], comparisonWorkBudget) >
                    GetCellSimilarity(sourceCellComparisonTexts[sourceIndex], targetCellComparisonTexts[targetIndex], comparisonWorkBudget)) {
                    AddDeletedTableCellFinding(sourceCells, tableIndex, tableDocumentOrder, sourceRowIndex, sourceIndex, result);
                    sourceIndex++;
                    continue;
                }

                string sourceText = GetCellText(sourceCells[sourceIndex]);
                string targetText = GetCellText(targetCells[targetIndex]);
                string sourceMatchText = GetCellMatchText(sourceCells[sourceIndex], sourcePart, options);
                string targetMatchText = GetCellMatchText(targetCells[targetIndex], targetPart, options);
                string sourceCellShape = GetCellShape(sourceCells[sourceIndex]);
                string targetCellShape = GetCellShape(targetCells[targetIndex]);

                if (!string.Equals(sourceMatchText, targetMatchText, StringComparison.Ordinal) ||
                    !string.Equals(sourceCellShape, targetCellShape, StringComparison.Ordinal)) {
                    result.Add(new WordComparisonFinding(
                        WordComparisonScope.TableCell,
                        WordComparisonChangeKind.Modified,
                        CellLocation(tableIndex, targetRowIndex, targetIndex),
                        sourceIndex,
                        targetIndex,
                        sourceText,
                        targetText,
                        AreComparisonTextEqual(sourceText, targetText, options) ? "Table cell structure changed." : "Table cell text changed."),
                        GetTableChildDocumentOrder(tableDocumentOrder, targetCells[targetIndex]));
                }

                sourceIndex++;
                targetIndex++;
            }

            while (targetIndex < targetEnd) {
                AddInsertedTableCellFinding(targetCells, tableIndex, tableDocumentOrder, targetRowIndex, targetIndex, result);
                targetIndex++;
            }

            while (sourceIndex < sourceEnd) {
                AddDeletedTableCellFinding(sourceCells, tableIndex, tableDocumentOrder, sourceRowIndex, sourceIndex, result);
                sourceIndex++;
            }
        }

        private static void AddInsertedTableCellFinding(IReadOnlyList<OfficeXmlNode> targetCells, int tableIndex, int tableDocumentOrder, int targetRowIndex, int cellIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.TableCell,
                WordComparisonChangeKind.Inserted,
                CellLocation(tableIndex, targetRowIndex, cellIndex),
                null,
                cellIndex,
                null,
                GetCellText(targetCells[cellIndex]),
                "Table cell inserted."),
                GetTableChildDocumentOrder(tableDocumentOrder, targetCells[cellIndex]));
        }

        private static void AddDeletedTableCellFinding(IReadOnlyList<OfficeXmlNode> sourceCells, int tableIndex, int tableDocumentOrder, int sourceRowIndex, int cellIndex, WordComparisonResult result) {
            result.Add(new WordComparisonFinding(
                WordComparisonScope.TableCell,
                WordComparisonChangeKind.Deleted,
                CellLocation(tableIndex, sourceRowIndex, cellIndex),
                cellIndex,
                null,
                GetCellText(sourceCells[cellIndex]),
                null,
                "Table cell deleted."),
                GetTableChildDocumentOrder(tableDocumentOrder, sourceCells[cellIndex]));
        }

        // Myers' linear-space diff from OfficeIMO.Core keeps the alignment exact for large documents; the greedy
        // pass is only a fallback when the diff work limit is reached.
        private static IReadOnlyList<MatchedIndexPair> FindMatchingIndexes<T>(IReadOnlyList<T> source, IReadOnlyList<T> target, IEqualityComparer<T> comparer) where T : notnull {
            var keyComparer = new FingerprintedValueComparer<T>(comparer);
            OfficeDiffResult<FingerprintedValue<T>> diff = OfficeDiff.Compare(
                FingerprintedValue<T>.Create(source, comparer),
                FingerprintedValue<T>.Create(target, comparer),
                keyComparer);
            if (diff.LimitReached) {
                return FindGreedyMatchingIndexes(source, target, comparer);
            }

            var matches = new List<MatchedIndexPair>(diff.EqualCount);
            foreach (OfficeDiffEdit edit in diff.Edits) {
                if (edit.Operation != OfficeDiffOperation.Equal) continue;
                for (int offset = 0; offset < edit.OldLength; offset++) {
                    matches.Add(new MatchedIndexPair(edit.OldStart + offset, edit.NewStart + offset));
                }
            }

            return matches;
        }

        /// <summary>A value with its hash code and independent fingerprint computed once; both must agree before the comparer is asked.</summary>
        private readonly struct FingerprintedValue<T> where T : notnull {
            private FingerprintedValue(T value, int hashCode, ulong fingerprint) {
                Value = value;
                HashCode = hashCode;
                Fingerprint = fingerprint;
            }

            internal T Value { get; }
            internal int HashCode { get; }
            internal ulong Fingerprint { get; }

            internal static FingerprintedValue<T>[] Create(IReadOnlyList<T> values, IEqualityComparer<T> comparer) {
                var result = new FingerprintedValue<T>[values.Count];
                for (int index = 0; index < values.Count; index++) {
                    result[index] = new FingerprintedValue<T>(values[index], comparer.GetHashCode(values[index]), GetComparisonFingerprint(values[index]));
                }

                return result;
            }
        }

        private sealed class FingerprintedValueComparer<T> : IEqualityComparer<FingerprintedValue<T>> where T : notnull {
            private readonly IEqualityComparer<T> _comparer;

            internal FingerprintedValueComparer(IEqualityComparer<T> comparer) {
                _comparer = comparer;
            }

            public bool Equals(FingerprintedValue<T> x, FingerprintedValue<T> y) =>
                x.HashCode == y.HashCode && x.Fingerprint == y.Fingerprint && _comparer.Equals(x.Value, y.Value);

            public int GetHashCode(FingerprintedValue<T> obj) => obj.HashCode;
        }

        private static IReadOnlyList<MatchedIndexPair> FindGreedyMatchingIndexes<T>(IReadOnlyList<T> source, IReadOnlyList<T> target, IEqualityComparer<T> comparer) where T : notnull {
            var matches = new List<MatchedIndexPair>();
            var targetIndexes = new Dictionary<T, Queue<int>>(comparer);
            for (int targetIndex = 0; targetIndex < target.Count; targetIndex++) {
                if (!targetIndexes.TryGetValue(target[targetIndex], out Queue<int>? indexes)) {
                    indexes = new Queue<int>();
                    targetIndexes.Add(target[targetIndex], indexes);
                }

                indexes.Enqueue(targetIndex);
            }

            int targetCursor = 0;

            for (int sourceIndex = 0; sourceIndex < source.Count && targetCursor < target.Count; sourceIndex++) {
                if (!targetIndexes.TryGetValue(source[sourceIndex], out Queue<int>? indexes)) {
                    continue;
                }

                while (indexes.Count > 0 && indexes.Peek() < targetCursor) {
                    indexes.Dequeue();
                }

                if (indexes.Count > 0) {
                    int targetIndex = indexes.Dequeue();
                    matches.Add(new MatchedIndexPair(sourceIndex, targetIndex));
                    targetCursor = targetIndex + 1;
                }
            }

            return matches;
        }

        private static List<TableSnapshot> GetTableSnapshots(WordPackageView document, WordComparisonOptions options) {
            var snapshots = new List<TableSnapshot>();
            WordPartView? mainPart = document.Main;
            AddTableSnapshots(snapshots, mainPart, BodyOf(mainPart), BodyPartKey, BodyPartOrderBase, options);

            if (mainPart != null) {
                int headerIndex = 0;
                foreach (KeyValuePair<WordPartView, string> headerPartKey in CreateOrderedHeaderPartKeys(mainPart)) {
                    AddTableSnapshots(snapshots, headerPartKey.Key, HeaderRoot(headerPartKey.Key), headerPartKey.Value, HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride), options);
                    headerIndex++;
                }

                int footerIndex = 0;
                foreach (KeyValuePair<WordPartView, string> footerPartKey in CreateOrderedFooterPartKeys(mainPart)) {
                    AddTableSnapshots(snapshots, footerPartKey.Key, FooterRoot(footerPartKey.Key), footerPartKey.Value, FooterPartOrderBase + (footerIndex * RelatedPartOrderStride), options);
                    footerIndex++;
                }

                List<OfficeXmlNode> footnotes = GetReferencedFootnotes(mainPart);
                for (int footnoteIndex = 0; footnoteIndex < footnotes.Count; footnoteIndex++) {
                    string noteId = GetNotePartKeyId(footnotes[footnoteIndex], footnoteIndex);
                    AddTableSnapshots(snapshots, FootnotesPart(mainPart), footnotes[footnoteIndex], FootnotePartKeyPrefix + noteId, FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride), options);
                }

                List<OfficeXmlNode> endnotes = GetReferencedEndnotes(mainPart);
                for (int endnoteIndex = 0; endnoteIndex < endnotes.Count; endnoteIndex++) {
                    string noteId = GetNotePartKeyId(endnotes[endnoteIndex], endnoteIndex);
                    AddTableSnapshots(snapshots, EndnotesPart(mainPart), endnotes[endnoteIndex], EndnotePartKeyPrefix + noteId, EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride), options);
                }
            }

            return snapshots;
        }

        private static void AddTableSnapshots(List<TableSnapshot> snapshots, WordPartView? part, OfficeXmlNode container, string partKey, int orderBase, WordComparisonOptions options) {
            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                if (!ordered.Element.Is(W, "tbl")) {
                    continue;
                }

                OfficeXmlNode table = ordered.Element;
                string text = GetTableText(table);
                string matchText = GetTableMatchText(table, part, options);
                snapshots.Add(new TableSnapshot(table, part, text, NormalizeComparisonText(text, options), matchText, GetTableMatchKey(partKey, table, part, options), partKey, ordered.DocumentOrder));
            }
        }

        private const int MaximumTableContainerDepth = 128;

        /// <summary>The logical rows of a table, through its content-control and custom-XML row containers.</summary>
        private static List<OfficeXmlNode> TableRows(OfficeXmlNode table) {
            var rows = new List<OfficeXmlNode>();
            CollectTableContent(table, "tr", rows, 0);
            return rows;
        }

        /// <summary>The logical cells of a row, through its content-control and custom-XML cell containers.</summary>
        private static List<OfficeXmlNode> RowCells(OfficeXmlNode row) {
            var cells = new List<OfficeXmlNode>();
            CollectTableContent(row, "tc", cells, 0);
            return cells;
        }

        private static void CollectTableContent(OfficeXmlNode parent, string itemName, List<OfficeXmlNode> items, int depth) {
            for (OfficeXmlNode child = parent.FirstChild; !child.IsNone; child = child.NextSibling) {
                if (child.Is(W, itemName)) {
                    // A cell is a boundary: its nested tables belong to that cell.
                    items.Add(child);
                    continue;
                }

                OfficeXmlNode content = child.Is(W, "sdt") ? child.Child(W, "sdtContent") : child.Is(W, "customXml") ? child : default;
                if (content.IsNone) {
                    continue;
                }

                if (depth >= MaximumTableContainerDepth) {
                    throw new InvalidDataException($"Table row or cell containers exceed the supported limit of {MaximumTableContainerDepth} levels.");
                }

                CollectTableContent(content, itemName, items, depth + 1);
            }
        }

        private static string GetTableText(OfficeXmlNode table) {
            return string.Join(TableRowSeparator, TableRows(table).Select(GetRowText).ToArray());
        }

        private static string GetRowText(OfficeXmlNode row) {
            return string.Join(" | ", RowCells(row).Select(GetCellText).ToArray());
        }

        private static string GetTableMatchText(OfficeXmlNode table, WordPartView? part, WordComparisonOptions options) {
            return string.Join(TableRowSeparator, TableRows(table).Select(row => GetRowMatchText(row, part, options)).ToArray());
        }

        private static string GetRowMatchText(OfficeXmlNode row, WordPartView? part, WordComparisonOptions options) {
            return string.Join(" | ", RowCells(row).Select(cell => GetCellMatchText(cell, part, options)).ToArray());
        }

        private static string GetTableMatchKey(string partKey, OfficeXmlNode table, WordPartView? part, WordComparisonOptions options) {
            List<OfficeXmlNode> rows = TableRows(table);
            return partKey + TableRowSeparator + string.Join(";", rows.Select(GetRowShape).ToArray()) + TableRowSeparator + string.Join(TableRowSeparator, rows.Select(row => GetRowMatchKey(row, part, options)).ToArray());
        }

        private static string GetRowMatchKey(OfficeXmlNode row, WordPartView? part, WordComparisonOptions options) {
            return GetRowShape(row) + TableRowSeparator + string.Join(TableRowSeparator, RowCells(row).Select(cell => EncodeMatchText(GetCellMatchText(cell, part, options))).ToArray());
        }

        private static string GetRowShape(OfficeXmlNode row) {
            List<OfficeXmlNode> cells = RowCells(row);
            return cells.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" +
                   string.Join(",", cells.Select(GetCellShape).ToArray());
        }

        private static string GetCellShape(OfficeXmlNode cell) {
            OfficeXmlNode properties = cell.Child(W, "tcPr");
            if (properties.IsNone) {
                return string.Empty;
            }

            string gridSpan = Int32Text(WVal(properties.Child(W, "gridSpan"))) ?? string.Empty;
            string horizontalMerge = GetMergeShapeValue(properties.Child(W, "hMerge"));
            string verticalMerge = GetMergeShapeValue(properties.Child(W, "vMerge"));
            return "span=" + gridSpan + ";h=" + horizontalMerge + ";v=" + verticalMerge;
        }

        /// <summary>
        /// The merge shape as the SDK path wrote it: "continue" without a value, and for a value the SDK enumeration's
        /// text, which named the type rather than the value.
        /// </summary>
        private static string GetMergeShapeValue(OfficeXmlNode merge) {
            if (merge.IsNone) {
                return string.Empty;
            }

            return EnumToken(WVal(merge), "restart", "continue") == null ? "continue" : "MergedCellValues { }";
        }

        private static IEnumerable<OfficeXmlNode> CellParagraphs(OfficeXmlNode cell) =>
            DescendantsW(cell, "p").Where(paragraph => AncestorW(paragraph, "tc") == cell);

        private static string GetCellText(OfficeXmlNode cell) {
            return string.Join(CellParagraphSeparator, CellParagraphs(cell).Select(GetParagraphText).ToArray());
        }

        private static string GetCellMatchText(OfficeXmlNode cell) {
            return GetCellMatchText(cell, null);
        }

        private static string GetCellMatchText(OfficeXmlNode cell, WordPartView? part) {
            return string.Join(CellParagraphSeparator, CellParagraphs(cell).Select(paragraph => GetParagraphMatchText(paragraph, part)).ToArray());
        }

        private static string GetCellMatchText(OfficeXmlNode cell, WordPartView? part, WordComparisonOptions options) {
            return string.Join(CellParagraphSeparator, CellParagraphs(cell).Select(paragraph => GetParagraphMatchText(paragraph, part, options)).ToArray());
        }

        private static string GetCellMatchKey(OfficeXmlNode cell, WordPartView? part) {
            return GetCellShape(cell) + CellParagraphSeparator + GetCellMatchText(cell, part);
        }

        private static string GetCellMatchKey(OfficeXmlNode cell, WordPartView? part, WordComparisonOptions options) {
            return GetCellShape(cell) + CellParagraphSeparator + GetCellMatchText(cell, part, options);
        }

        private static string EncodeMatchText(string value) {
            return value.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + value;
        }

        private static double GetRowSimilarity(
            OfficeXmlNode source,
            string sourceComparisonText,
            OfficeXmlNode target,
            string targetComparisonText,
            ComparisonWorkBudget comparisonWorkBudget) {
            if (RowCells(source).Count != RowCells(target).Count) {
                return 0;
            }

            return GetContainmentAwareTextSimilarity(sourceComparisonText, targetComparisonText, comparisonWorkBudget);
        }

        private static double GetRowAlignmentPrefilterSimilarity(
            OfficeXmlNode source,
            string sourceComparisonText,
            OfficeXmlNode target,
            string targetComparisonText,
            ComparisonWorkBudget comparisonWorkBudget) {
            if (RowCells(source).Count != RowCells(target).Count) {
                return 0;
            }

            return GetBoundedTextSimilarity(sourceComparisonText, targetComparisonText, comparisonWorkBudget);
        }

        private static double GetTableSimilarity(TableSnapshot source, TableSnapshot target, ComparisonWorkBudget comparisonWorkBudget) {
            if (!string.Equals(source.PartKey, target.PartKey, StringComparison.Ordinal)) {
                return 0;
            }

            return GetTextSimilarity(source.ComparisonText, target.ComparisonText, comparisonWorkBudget);
        }

        private static double GetCellSimilarity(string sourceComparisonText, string targetComparisonText, ComparisonWorkBudget comparisonWorkBudget) {
            return GetTextSimilarity(sourceComparisonText, targetComparisonText, comparisonWorkBudget);
        }

        /// <summary>
        /// A path of element names and ordinals among same-named siblings, as the SDK path built it from typed classes;
        /// siblings share a parent, so their names tell the same classes apart.
        /// </summary>
        private static string GetStableElementPath(OfficeXmlNode element) => GetStableElementPathRelativeTo(element, default);

        private static string GetStableElementPathRelativeTo(OfficeXmlNode element, OfficeXmlNode root) {
            var segments = new Stack<string>();
            OfficeXmlNode current = element;
            while (!current.IsNone && !current.Parent.IsNone && current != root) {
                OfficeXmlNode parent = current.Parent;
                string namespaceUri = current.NamespaceUri;
                string localName = current.LocalName;
                int ordinal = 0;
                for (OfficeXmlNode sibling = parent.FirstChild; !sibling.IsNone && sibling != current; sibling = sibling.NextSibling) {
                    if (string.Equals(sibling.LocalName, localName, StringComparison.Ordinal) &&
                        string.Equals(sibling.NamespaceUri, namespaceUri, StringComparison.Ordinal)) {
                        ordinal++;
                    }
                }

                segments.Push("{" + namespaceUri + "}" + localName + "[" + ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]");
                current = parent;
            }

            return string.Join("/", segments.ToArray());
        }

        private static string TableLocation(int tableIndex) {
            return "table[" + tableIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
        }

        private static string RowLocation(int tableIndex, int rowIndex) {
            return TableLocation(tableIndex) + "/row[" + rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
        }

        private static string CellLocation(int tableIndex, int rowIndex, int cellIndex) {
            return RowLocation(tableIndex, rowIndex) + "/cell[" + cellIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
        }

        private static int GetTableChildDocumentOrder(int tableDocumentOrder, OfficeXmlNode child) {
            OfficeXmlNode table = child.Is(W, "tbl") ? child : AncestorW(child, "tbl");
            if (table.IsNone) {
                return tableDocumentOrder;
            }

            int offset = 1;
            foreach (OfficeXmlNode descendant in Descendants(table)) {
                if (descendant == child) {
                    return tableDocumentOrder + offset;
                }

                offset++;
            }

            return tableDocumentOrder;
        }

        private readonly struct OrderedElement {
            internal OrderedElement(OfficeXmlNode element, int documentOrder) {
                Element = element;
                DocumentOrder = documentOrder;
            }

            internal OfficeXmlNode Element { get; }

            internal int DocumentOrder { get; }
        }

        private readonly struct MatchedIndexPair {
            internal MatchedIndexPair(int sourceIndex, int targetIndex) {
                SourceIndex = sourceIndex;
                TargetIndex = targetIndex;
            }

            internal int SourceIndex { get; }

            internal int TargetIndex { get; }
        }

        private sealed class TableSnapshot {
            internal TableSnapshot(OfficeXmlNode table, WordPartView? part, string text, string comparisonText, string matchText, string matchKey, string partKey, int documentOrder) {
                Table = table;
                Part = part;
                Text = text;
                ComparisonText = comparisonText;
                MatchText = matchText;
                MatchKey = matchKey;
                PartKey = partKey;
                DocumentOrder = documentOrder;
            }

            internal OfficeXmlNode Table { get; }

            internal WordPartView? Part { get; }

            internal string Text { get; }

            internal string ComparisonText { get; }

            internal string MatchText { get; }

            internal string MatchKey { get; }

            internal string PartKey { get; }

            internal int DocumentOrder { get; }
        }

    }
}
