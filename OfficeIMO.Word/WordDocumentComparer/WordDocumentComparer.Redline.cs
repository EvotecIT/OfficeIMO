using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        /// <summary>
        /// Compares two documents and writes a DOCX redline document using the configured redline mode.
        /// </summary>
        /// <param name="sourcePath">Path to the original document.</param>
        /// <param name="targetPath">Path to the modified document.</param>
        /// <param name="outputPath">Path where the redline DOCX should be written.</param>
        /// <param name="options">Optional redline and comparison settings.</param>
        /// <returns>The structured comparison result used to generate the redline document.</returns>
        public static WordComparisonResult CreateRedlineDocument(string sourcePath, string targetPath, string outputPath, WordComparisonRedlineOptions? options = null) {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (string.IsNullOrEmpty(targetPath)) throw new ArgumentNullException(nameof(targetPath));
            if (string.IsNullOrEmpty(outputPath)) throw new ArgumentNullException(nameof(outputPath));

            options ??= new WordComparisonRedlineOptions();
            WordComparisonResult result = CompareStructure(sourcePath, targetPath, options.ComparisonOptions);

            EnsureOutputDirectory(outputPath);

            if (options.Mode == WordComparisonRedlineMode.InPlaceTarget) {
                CreateInPlaceTargetRedlineDocument(sourcePath, targetPath, outputPath, result, options);
                return result;
            }

            using WordDocument document = WordDocument.Create(outputPath);
            document.AddParagraph("Word Comparison Redline").SetStyle(WordParagraphStyles.Heading1);
            document.AddParagraph(WordComparisonReportWriter.ToTextSummary(result));

            if (options.IncludeSummary) {
                AppendRedlineSummary(document, result);
            }

            if (options.IncludeFindingsTable) {
                AppendRedlineFindingsTable(document, result);
            }

            if (ShouldAppendTrackedRedlineFindings(options)) {
                AppendTrackedRedlineFindings(document, result, options);
            }

            document.Save(false);
            return result;
        }

        private static void EnsureOutputDirectory(string outputPath) {
            string? directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory)) {
                Directory.CreateDirectory(directory);
            }
        }

        private static void CreateInPlaceTargetRedlineDocument(string sourcePath, string targetPath, string outputPath, WordComparisonResult result, WordComparisonRedlineOptions options) {
            if (string.Equals(Path.GetFullPath(targetPath), Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidOperationException("In-place target redline output must be written to a different path than the target document.");
            }

            File.Copy(targetPath, outputPath, overwrite: true);
            using WordDocument sourceDocument = WordDocument.Load(sourcePath, readOnly: true);
            using WordDocument document = WordDocument.Load(outputPath);
            HashSet<int> rewrittenParagraphs = ApplyParagraphFindings(sourceDocument._wordprocessingDocument, document._wordprocessingDocument, result, options);
            ApplyRunFindings(sourceDocument._wordprocessingDocument, document._wordprocessingDocument, result, options, rewrittenParagraphs);
            ApplyContentControlFindings(document._wordprocessingDocument, result, options);
            ApplyImageFindings(sourceDocument._wordprocessingDocument, document._wordprocessingDocument, result, options);

            ApplyTableFindings(sourceDocument._wordprocessingDocument, document._wordprocessingDocument, result, options);
            ApplyTableCellFindings(sourceDocument._wordprocessingDocument, document._wordprocessingDocument, result, options);
            ApplyTableRowFindings(sourceDocument._wordprocessingDocument, document._wordprocessingDocument, result, options);

            document.Save(false);
        }

        private static bool HasTrackedText(WordComparisonFinding finding) {
            return !string.IsNullOrEmpty(finding.SourceText) || !string.IsNullOrEmpty(finding.TargetText);
        }

        private static void RewriteParagraphWithTrackedText(Paragraph paragraph, string? sourceText, string? targetText, WordComparisonRedlineOptions options) {
            foreach (OpenXmlElement child in paragraph.ChildElements.Where(child => child is not ParagraphProperties && !ShouldPreserveParagraphRedlineChild(child)).ToList()) {
                child.Remove();
            }

            OpenXmlElement? insertionPoint = paragraph.ChildElements.FirstOrDefault(child => child is not ParagraphProperties);
            if (!string.IsNullOrEmpty(sourceText)) {
                InsertParagraphRedlineRun(paragraph, insertionPoint, CreateDeletedRun(sourceText!, options));
            }

            if (!string.IsNullOrEmpty(targetText)) {
                InsertParagraphRedlineRun(paragraph, insertionPoint, CreateInsertedRun(targetText!, options));
            }
        }

        private static bool ShouldPreserveParagraphRedlineChild(OpenXmlElement child) {
            if (child is not Run run) {
                return true;
            }

            return run.ChildElements.Any(runChild => runChild is not RunProperties && runChild is not Text);
        }

        private static void InsertParagraphRedlineRun(Paragraph paragraph, OpenXmlElement? insertionPoint, OpenXmlElement run) {
            if (insertionPoint != null) {
                paragraph.InsertBefore(run, insertionPoint);
                return;
            }

            paragraph.Append(run);
        }

        private static void ApplyRunFinding(Paragraph paragraph, int runIndex, WordComparisonFinding finding, WordComparisonRedlineOptions options) {
            List<Run> runs = paragraph.Descendants<Run>()
                .Where(run => run.Ancestors<Paragraph>().FirstOrDefault() == paragraph)
                .ToList();
            if (runs.Count == 0) {
                RewriteParagraphWithTrackedText(paragraph, finding.SourceText, finding.TargetText, options);
                return;
            }

            int boundedRunIndex = Math.Max(0, Math.Min(runIndex, runs.Count - 1));
            Run targetRun = runs[boundedRunIndex];
            OpenXmlElement parent = targetRun.Parent ?? paragraph;

            if (finding.ChangeKind != WordComparisonChangeKind.Inserted && !string.IsNullOrEmpty(finding.SourceText)) {
                parent.InsertBefore(CreateDeletedRun(finding.SourceText!, options), targetRun);
            }

            if (finding.ChangeKind != WordComparisonChangeKind.Deleted && !string.IsNullOrEmpty(finding.TargetText)) {
                parent.InsertBefore(CreateInsertedRun(finding.TargetText!, options), targetRun);
            }

            if (finding.ChangeKind != WordComparisonChangeKind.Deleted) {
                targetRun.Remove();
            }
        }

        private static void ApplyTableCellFindings(WordprocessingDocument sourceDocument, WordprocessingDocument targetDocument, WordComparisonResult result, WordComparisonRedlineOptions options) {
            List<RedlineTableEntry> sourceTables = GetRedlineTableEntries(sourceDocument);
            List<RedlineTableEntry> targetTables = GetRedlineTableEntries(targetDocument);
            HashSet<string> nestedTableParentCellKeys = GetNestedTableParentCellKeys(sourceTables, targetTables, result, options);
            var rewrittenCells = new HashSet<string>(StringComparer.Ordinal);
            foreach (WordComparisonFinding finding in result.Findings) {
                if (!ShouldTrackFinding(finding, options) ||
                    finding.Scope != WordComparisonScope.TableCell ||
                    !TryParseTableCellLocation(finding.Location, out int tableIndex, out int rowIndex, out int cellIndex) ||
                    tableIndex < 0 ||
                    tableIndex >= targetTables.Count ||
                    rowIndex < 0 ||
                    !HasTrackedText(finding)) {
                    continue;
                }

                Table table = targetTables[tableIndex].Table;
                List<TableRow> rows = table.Elements<TableRow>().ToList();
                if (rowIndex >= rows.Count) {
                    continue;
                }

                TableRow row = rows[rowIndex];
                List<TableCell> cells = row.Elements<TableCell>().ToList();
                string cellKey = CreateTableCellKey(tableIndex, rowIndex, cellIndex);
                if (rewrittenCells.Contains(cellKey) || nestedTableParentCellKeys.Contains(cellKey)) {
                    continue;
                }

                switch (finding.ChangeKind) {
                    case WordComparisonChangeKind.Modified:
                    case WordComparisonChangeKind.Inserted:
                        if (cellIndex >= 0 && cellIndex < cells.Count) {
                            RewriteCellWithTrackedText(cells[cellIndex], finding.SourceText, finding.TargetText, options);
                            RemoveEmptyWordColorAttributes(targetTables[tableIndex].Table);
                            rewrittenCells.Add(cellKey);
                        }

                        break;
                    case WordComparisonChangeKind.Deleted:
                        if (!string.IsNullOrEmpty(finding.SourceText)) {
                            InsertDeletedCell(row, cells, cellIndex, finding.SourceText!, options);
                            RemoveEmptyWordColorAttributes(targetTables[tableIndex].Table);
                            rewrittenCells.Add(cellKey);
                        }

                        break;
                }
            }
        }

        private static HashSet<string> GetNestedTableParentCellKeys(IReadOnlyList<RedlineTableEntry> sourceTables, IReadOnlyList<RedlineTableEntry> targetTables, WordComparisonResult result, WordComparisonRedlineOptions options) {
            var cellKeys = new HashSet<string>(StringComparer.Ordinal);
            foreach (WordComparisonFinding finding in result.Findings) {
                if (!ShouldTrackFinding(finding, options) ||
                    finding.Scope != WordComparisonScope.Table ||
                    !TryParseTableLocation(finding.Location, out int tableIndex)) {
                    continue;
                }

                IReadOnlyList<RedlineTableEntry> tables = finding.ChangeKind == WordComparisonChangeKind.Deleted ? sourceTables : targetTables;
                if (tableIndex < 0 || tableIndex >= tables.Count) {
                    continue;
                }

                if (TryGetNestedTablePlacement(tables, tables[tableIndex].Table, out NestedTablePlacement placement)) {
                    cellKeys.Add(CreateTableCellKey(placement.ParentTableIndex, placement.RowIndex, placement.CellIndex));
                }
            }

            return cellKeys;
        }

        private static void ApplyTableRowFindings(WordprocessingDocument sourceDocument, WordprocessingDocument targetDocument, WordComparisonResult result, WordComparisonRedlineOptions options) {
            List<RedlineTableEntry> sourceTables = GetRedlineTableEntries(sourceDocument);
            List<RedlineTableEntry> targetTables = GetRedlineTableEntries(targetDocument);
            var rewrittenRows = new HashSet<string>(StringComparer.Ordinal);

            foreach (WordComparisonFinding finding in result.Findings) {
                if (!ShouldTrackFinding(finding, options) ||
                    finding.Scope != WordComparisonScope.TableRow ||
                    !TryParseTableRowLocation(finding.Location, out int tableIndex, out int rowIndex) ||
                    tableIndex < 0 ||
                    !HasTrackedText(finding)) {
                    continue;
                }

                string rowKey = tableIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "/" +
                                rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (rewrittenRows.Contains(rowKey)) {
                    continue;
                }

                switch (finding.ChangeKind) {
                    case WordComparisonChangeKind.Inserted:
                        if (tableIndex < targetTables.Count) {
                            List<TableRow> targetRows = targetTables[tableIndex].Table.Elements<TableRow>().ToList();
                            if (rowIndex >= 0 && rowIndex < targetRows.Count) {
                                RewriteRowWithTrackedText(targetRows[rowIndex], trackInserted: true, options);
                                rewrittenRows.Add(rowKey);
                            }
                        }

                        break;
                    case WordComparisonChangeKind.Deleted:
                        if (tableIndex < sourceTables.Count && tableIndex < targetTables.Count) {
                            List<TableRow> sourceRows = sourceTables[tableIndex].Table.Elements<TableRow>().ToList();
                            if (rowIndex >= 0 && rowIndex < sourceRows.Count) {
                                InsertDeletedRow(targetTables[tableIndex].Table, sourceRows[rowIndex], rowIndex, options);
                                rewrittenRows.Add(rowKey);
                            }
                        }

                        break;
                }
            }
        }

        private static List<RedlineTableEntry> GetRedlineTableEntries(WordprocessingDocument document) {
            var entries = new List<RedlineTableEntry>();
            MainDocumentPart? mainPart = document.MainDocumentPart;
            AddRedlineTableEntries(entries, BodyPartKey, mainPart?.Document?.Body, BodyPartOrderBase);

            if (mainPart == null) {
                return entries;
            }

            int headerIndex = 0;
            foreach (KeyValuePair<HeaderPart, string> headerPartKey in CreateOrderedHeaderPartKeys(mainPart)) {
                AddRedlineTableEntries(entries, headerPartKey.Value, headerPartKey.Key.Header, HeaderPartOrderBase + (headerIndex * RelatedPartOrderStride));
                headerIndex++;
            }

            int footerIndex = 0;
            foreach (KeyValuePair<FooterPart, string> footerPartKey in CreateOrderedFooterPartKeys(mainPart)) {
                AddRedlineTableEntries(entries, footerPartKey.Value, footerPartKey.Key.Footer, FooterPartOrderBase + (footerIndex * RelatedPartOrderStride));
                footerIndex++;
            }

            List<Footnote> footnotes = GetReferencedFootnotes(mainPart);
            for (int footnoteIndex = 0; footnoteIndex < footnotes.Count; footnoteIndex++) {
                string noteId = footnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                AddRedlineTableEntries(entries, FootnotePartKeyPrefix + noteId, footnotes[footnoteIndex], FootnotePartOrderBase + (footnoteIndex * RelatedPartOrderStride));
            }

            List<Endnote> endnotes = GetReferencedEndnotes(mainPart);
            for (int endnoteIndex = 0; endnoteIndex < endnotes.Count; endnoteIndex++) {
                string noteId = endnoteIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                AddRedlineTableEntries(entries, EndnotePartKeyPrefix + noteId, endnotes[endnoteIndex], EndnotePartOrderBase + (endnoteIndex * RelatedPartOrderStride));
            }

            return entries;
        }

        private static void AddRedlineTableEntries(List<RedlineTableEntry> entries, string partKey, OpenXmlCompositeElement? container, int orderBase) {
            if (container == null) {
                return;
            }

            foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(container, orderBase)) {
                if (ordered.Element is Table table) {
                    entries.Add(new RedlineTableEntry(partKey, container, table));
                }
            }
        }

        private static void RewriteCellWithTrackedText(TableCell cell, string? sourceText, string? targetText, WordComparisonRedlineOptions options) {
            TableCellProperties? properties = cell.GetFirstChild<TableCellProperties>()?.CloneNode(true) as TableCellProperties;
            cell.RemoveAllChildren();
            if (properties != null) {
                cell.Append(properties);
            }

            AppendCellTrackedParagraphs(cell, sourceText, targetText, options);
            EnsureCellHasParagraph(cell);
            RemoveEmptyWordColorAttributes(cell);
        }

        private static void InsertDeletedCell(TableRow row, IReadOnlyList<TableCell> existingCells, int cellIndex, string sourceText, WordComparisonRedlineOptions options) {
            var deletedCell = new TableCell();
            AppendCellTrackedParagraphs(deletedCell, sourceText, null, options);
            EnsureCellHasParagraph(deletedCell);

            if (cellIndex >= 0 && cellIndex < existingCells.Count) {
                existingCells[cellIndex].InsertBeforeSelf(deletedCell);
            } else {
                row.Append(deletedCell);
            }
        }

        private static void RewriteRowWithTrackedText(TableRow row, bool trackInserted, WordComparisonRedlineOptions options) {
            foreach (TableCell cell in row.Elements<TableCell>()) {
                string cellText = GetOpenXmlCellText(cell);
                if (trackInserted) {
                    RewriteCellWithTrackedText(cell, null, cellText, options);
                } else {
                    RewriteCellWithTrackedText(cell, cellText, null, options);
                }
            }
        }

        private static void InsertDeletedRow(Table targetTable, TableRow sourceRow, int rowIndex, WordComparisonRedlineOptions options) {
            var deletedRow = (TableRow)sourceRow.CloneNode(true);
            RewriteRowWithTrackedText(deletedRow, trackInserted: false, options);

            List<TableRow> targetRows = targetTable.Elements<TableRow>().ToList();
            if (rowIndex >= 0 && rowIndex < targetRows.Count) {
                targetRows[rowIndex].InsertBeforeSelf(deletedRow);
            } else {
                targetTable.Append(deletedRow);
            }
        }

        private static void AppendCellTrackedParagraphs(TableCell cell, string? sourceText, string? targetText, WordComparisonRedlineOptions options) {
            string[] sourceParagraphs = SplitCellParagraphText(sourceText);
            string[] targetParagraphs = SplitCellParagraphText(targetText);
            int paragraphCount = Math.Max(sourceParagraphs.Length, targetParagraphs.Length);
            for (int index = 0; index < paragraphCount; index++) {
                string? sourceParagraph = index < sourceParagraphs.Length ? sourceParagraphs[index] : null;
                string? targetParagraph = index < targetParagraphs.Length ? targetParagraphs[index] : null;
                var paragraph = new Paragraph();
                if (!string.IsNullOrEmpty(sourceParagraph)) {
                    paragraph.Append(CreateDeletedRun(sourceParagraph!, options));
                }

                if (!string.IsNullOrEmpty(targetParagraph)) {
                    paragraph.Append(CreateInsertedRun(targetParagraph!, options));
                }

                cell.Append(paragraph);
            }
        }

        private static string[] SplitCellParagraphText(string? text) {
            if (string.IsNullOrEmpty(text)) {
                return Array.Empty<string>();
            }

            return text!.Split(new[] { CellParagraphSeparator }, StringSplitOptions.None);
        }

        private static string GetOpenXmlCellText(TableCell cell) {
            return string.Join(
                CellParagraphSeparator,
                cell.Descendants<Paragraph>()
                    .Where(paragraph => ReferenceEquals(paragraph.Ancestors<TableCell>().FirstOrDefault(), cell))
                    .Select(paragraph => paragraph.InnerText)
                    .ToArray());
        }

        private static void EnsureCellHasParagraph(TableCell cell) {
            if (!cell.Elements<Paragraph>().Any()) {
                cell.Append(new Paragraph());
            }
        }

        private static bool TryParseTableCellLocation(string location, out int tableIndex, out int rowIndex, out int cellIndex) {
            tableIndex = -1;
            rowIndex = -1;
            cellIndex = -1;
            const string tablePrefix = "table[";
            const string rowToken = "]/row[";
            const string cellToken = "]/cell[";
            if (!location.StartsWith(tablePrefix, StringComparison.Ordinal)) {
                return false;
            }

            int rowTokenIndex = location.IndexOf(rowToken, StringComparison.Ordinal);
            int cellTokenIndex = location.IndexOf(cellToken, StringComparison.Ordinal);
            if (rowTokenIndex < 0 || cellTokenIndex < 0 || cellTokenIndex <= rowTokenIndex || !location.EndsWith("]", StringComparison.Ordinal)) {
                return false;
            }

            string tableText = location.Substring(tablePrefix.Length, rowTokenIndex - tablePrefix.Length);
            string rowText = location.Substring(rowTokenIndex + rowToken.Length, cellTokenIndex - rowTokenIndex - rowToken.Length);
            string cellText = location.Substring(cellTokenIndex + cellToken.Length, location.Length - cellTokenIndex - cellToken.Length - 1);
            return int.TryParse(tableText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out tableIndex) &&
                   int.TryParse(rowText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out rowIndex) &&
                   int.TryParse(cellText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out cellIndex);
        }

        private static bool TryParseTableRowLocation(string location, out int tableIndex, out int rowIndex) {
            tableIndex = -1;
            rowIndex = -1;
            const string tablePrefix = "table[";
            const string rowToken = "]/row[";
            if (!location.StartsWith(tablePrefix, StringComparison.Ordinal)) {
                return false;
            }

            int rowTokenIndex = location.IndexOf(rowToken, StringComparison.Ordinal);
            if (rowTokenIndex < 0 || !location.EndsWith("]", StringComparison.Ordinal)) {
                return false;
            }

            string tableText = location.Substring(tablePrefix.Length, rowTokenIndex - tablePrefix.Length);
            string rowText = location.Substring(rowTokenIndex + rowToken.Length, location.Length - rowTokenIndex - rowToken.Length - 1);
            return int.TryParse(tableText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out tableIndex) &&
                   int.TryParse(rowText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out rowIndex);
        }

        private static string CreateTableCellKey(int tableIndex, int rowIndex, int cellIndex) {
            return tableIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "/" +
                   rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "/" +
                   cellIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private sealed class RedlineTableEntry {
            internal RedlineTableEntry(string partKey, OpenXmlCompositeElement container, Table table) {
                PartKey = partKey;
                Container = container;
                Table = table;
            }

            internal string PartKey { get; }

            internal OpenXmlCompositeElement Container { get; }

            internal Table Table { get; }
        }

        private static InsertedRun CreateInsertedRun(string text, WordComparisonRedlineOptions options) {
            var run = new Run();
            run.RsidRunAddition = WordHeadersAndFooters.GenerateRsid();
            run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            var inserted = new InsertedRun {
                Author = options.Author,
                Date = options.DateTime ?? DateTime.Now,
                Id = WordHeadersAndFooters.GenerateRevisionId()
            };
            inserted.Append(run);
            return inserted;
        }

        private static DeletedRun CreateDeletedRun(string text, WordComparisonRedlineOptions options) {
            var run = new Run();
            run.RsidRunDeletion = WordHeadersAndFooters.GenerateRsid();
            run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            var deleted = new DeletedRun {
                Author = options.Author,
                Date = options.DateTime ?? DateTime.Now,
                Id = WordHeadersAndFooters.GenerateRevisionId()
            };
            deleted.Append(run);
            return deleted;
        }

        private static void AppendRedlineSummary(WordDocument document, WordComparisonResult result) {
            document.AddParagraph("Summary").SetStyle(WordParagraphStyles.Heading2);
            WordTable table = document.AddTable(4, 2);
            SetCellText(table, 0, 0, "Metric");
            SetCellText(table, 0, 1, "Value");
            SetCellText(table, 1, 0, "Source");
            SetCellText(table, 1, 1, result.SourcePath);
            SetCellText(table, 2, 0, "Target");
            SetCellText(table, 2, 1, result.TargetPath);
            SetCellText(table, 3, 0, "Findings");
            SetCellText(table, 3, 1, result.Findings.Count.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        private static void AppendRedlineFindingsTable(WordDocument document, WordComparisonResult result) {
            document.AddParagraph("Findings").SetStyle(WordParagraphStyles.Heading2);
            WordTable table = document.AddTable(result.Findings.Count + 1, 6);
            SetCellText(table, 0, 0, "#");
            SetCellText(table, 0, 1, "Scope");
            SetCellText(table, 0, 2, "Change");
            SetCellText(table, 0, 3, "Location");
            SetCellText(table, 0, 4, "Source");
            SetCellText(table, 0, 5, "Target");

            for (int i = 0; i < result.Findings.Count; i++) {
                WordComparisonFinding finding = result.Findings[i];
                int row = i + 1;
                SetCellText(table, row, 0, i.ToString(System.Globalization.CultureInfo.InvariantCulture));
                SetCellText(table, row, 1, finding.Scope.ToString());
                SetCellText(table, row, 2, finding.ChangeKind.ToString());
                SetCellText(table, row, 3, finding.Location);
                SetCellText(table, row, 4, finding.SourceText ?? string.Empty);
                SetCellText(table, row, 5, finding.TargetText ?? string.Empty);
            }
        }

        private static void AppendTrackedRedlineFindings(WordDocument document, WordComparisonResult result, WordComparisonRedlineOptions options) {
            document.AddParagraph("Tracked Changes").SetStyle(WordParagraphStyles.Heading2);
            if (result.Findings.Count == 0) {
                document.AddParagraph("No tracked changes were generated because no structural differences were detected.");
                return;
            }

            bool wroteAnyRevision = false;
            foreach (WordComparisonFinding finding in result.Findings) {
                document.AddParagraph(finding.Location + " - " + finding.Scope + " " + finding.ChangeKind);
                WordParagraph paragraph = document.AddParagraph();
                bool wroteRevision = false;

                if (ShouldTrackFinding(finding, options) && ShouldEmitDeletedText(finding)) {
                    paragraph.AddDeletedText(finding.SourceText!, options.Author, options.DateTime);
                    wroteRevision = true;
                }

                if (ShouldTrackFinding(finding, options) && ShouldEmitInsertedText(finding)) {
                    paragraph.AddInsertedText(finding.TargetText!, options.Author, options.DateTime);
                    wroteRevision = true;
                }

                if (!wroteRevision) {
                    paragraph.SetText(finding.Message);
                } else {
                    wroteAnyRevision = true;
                }
            }

            if (!wroteAnyRevision) {
                document.AddParagraph("No tracked changes were generated because the selected redline policy kept all findings report-only.");
            }
        }

        private static bool ShouldAppendTrackedRedlineFindings(WordComparisonRedlineOptions options) {
            return options.TrackTextFindings ||
                   options.TrackFeatureFindings ||
                   options.TrackReviewFindings ||
                   options.TrackFormattingFindings;
        }

        private static bool ShouldTrackFinding(WordComparisonFinding finding, WordComparisonRedlineOptions options) {
            if (!options.TrackFeatureFindings && IsFeatureFinding(finding)) {
                return false;
            }

            if (!options.TrackReviewFindings && IsReviewFinding(finding)) {
                return false;
            }

            if (!options.TrackFormattingFindings && IsFormattingFinding(finding)) {
                return false;
            }

            if (!options.TrackTextFindings && IsTextFinding(finding)) {
                return false;
            }

            return true;
        }

        private static bool IsFeatureFinding(WordComparisonFinding finding) {
            return finding.Scope is WordComparisonScope.Field
                or WordComparisonScope.ContentControl
                or WordComparisonScope.Bookmark
                or WordComparisonScope.Hyperlink
                or WordComparisonScope.List
                or WordComparisonScope.Image;
        }

        private static bool IsReviewFinding(WordComparisonFinding finding) {
            return finding.Scope is WordComparisonScope.Comment or WordComparisonScope.Revision;
        }

        private static bool IsFormattingFinding(WordComparisonFinding finding) {
            if (finding.Message.IndexOf("formatting", StringComparison.OrdinalIgnoreCase) >= 0) {
                return true;
            }

            string sourceText = finding.SourceText ?? string.Empty;
            string targetText = finding.TargetText ?? string.Empty;
            return ContainsReviewFormattingKind(sourceText) || ContainsReviewFormattingKind(targetText);
        }

        private static bool IsTextFinding(WordComparisonFinding finding) {
            return HasTrackedText(finding)
                && !IsFeatureFinding(finding)
                && !IsReviewFinding(finding)
                && !IsFormattingFinding(finding);
        }

        private static bool ContainsReviewFormattingKind(string value) {
            return value.IndexOf("Formatting", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool ShouldEmitDeletedText(WordComparisonFinding finding) {
            return finding.ChangeKind != WordComparisonChangeKind.Inserted && !string.IsNullOrEmpty(finding.SourceText);
        }

        private static bool ShouldEmitInsertedText(WordComparisonFinding finding) {
            return finding.ChangeKind != WordComparisonChangeKind.Deleted && !string.IsNullOrEmpty(finding.TargetText);
        }

        private static void SetCellText(WordTable table, int row, int column, string text) {
            table.Rows[row].Cells[column].Paragraphs[0].SetText(text);
        }
    }
}
